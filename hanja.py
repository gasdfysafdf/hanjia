"""한자 연습장 웹앱 (Streamlit)

- 한자 파일(csv / xlsx / txt) 업로드 -> 문제 자동 생성 ('급수/단원/분류' 열이 있으면 범위 선택 가능)
- 문제 유형: 뜻·음 쓰기(채팅 입력), 한자 쓰기(캔버스 손글씨), 객관식 2종
- 플래시카드, 복습 주기(간격 반복), 오늘의 목표, 날짜별 기록
- 오답 노트, 오답만 다시 풀기, 통계, 암기 모드, 시험지(PDF), 진행 상황 저장/불러오기
"""
import base64
import copy
import glob
import html as htmllib
import inspect
import io
import json
import math
import os
import random
import re
import tempfile
import time
import traceback
import urllib.request
from datetime import date, datetime, timedelta, timezone

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image, ImageDraw, ImageFont

CANVAS_ERR = ""
try:
    from streamlit_drawable_canvas import st_canvas

    HAS_CANVAS = True
except Exception as _e:  # 라이브러리가 없어도 뜻·음 문제는 풀 수 있게
    HAS_CANVAS = False
    CANVAS_ERR = f"{type(_e).__name__}: {_e}"

st.set_page_config(page_title="한자 연습장", page_icon="📖", layout="wide", initial_sidebar_state="collapsed")

# ----------------------------------------------------------------------
# 폰트 (손글씨 비교용 정답 글자를 그리려면 한자가 들어 있는 폰트가 필요)
# ----------------------------------------------------------------------
FONT_PATTERNS = [
    os.environ.get("FONT_PATH", ""),
    "fonts/*.ttf",
    "fonts/*.otf",
    "fonts/*.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "C:/Windows/Fonts/malgun.ttf",
    "C:/Windows/Fonts/batang.ttc",
    "/System/Library/Fonts/AppleSDGothicNeo.ttc",
    "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
]


def load_font(path, size):
    if path.lower().endswith(".ttc"):
        try:
            return ImageFont.truetype(path, size, index=1)  # Noto CJK: 1 = KR
        except Exception:
            return ImageFont.truetype(path, size, index=0)
    return ImageFont.truetype(path, size)


FONT_URLS = [
    "https://raw.githubusercontent.com/notofonts/noto-cjk/main/Sans/SubsetOTF/KR/NotoSansKR-Regular.otf",
    "https://cdn.jsdelivr.net/gh/notofonts/noto-cjk@main/Sans/SubsetOTF/KR/NotoSansKR-Regular.otf",  # 예비 주소
]


@st.cache_resource(show_spinner=False)
def _font_state():
    """폰트 위치와 '내려받기 실패 시각'을 화면이 다시 그려져도 기억해요 (실행마다 새로 초기화되지 않게)"""
    return {"path": None, "source": "", "fail_at": 0.0}


_FONT = _font_state()
FONT_RETRY_SEC = 120  # 내려받기에 실패하면 이 시간 동안은 다시 시도하지 않아요


def _system_font():
    pats = list(FONT_PATTERNS) + [
        "/usr/share/fonts/**/*CJK*",
        "/usr/share/fonts/**/*Noto*KR*",
        "/usr/share/fonts/**/*Nanum*",
    ]
    for pat in pats:
        if not pat:
            continue
        for p in sorted(glob.glob(pat, recursive=True)):
            if not p.lower().endswith((".ttf", ".otf", ".ttc")):
                continue
            try:
                load_font(p, 40)
                return p
            except Exception:
                continue
    return None


def _download_font():
    """시스템에 한자 폰트가 없을 때 Noto Sans KR(약 4.6MB)을 한 번 내려받아 사용"""
    path = os.path.join(tempfile.gettempdir(), "hanja_font", "NotoSansKR-Regular.otf")
    try:
        if os.path.exists(path) and os.path.getsize(path) > 1_000_000:
            load_font(path, 40)
            return path
    except Exception:
        pass
    os.makedirs(os.path.dirname(path), exist_ok=True)
    for url in FONT_URLS:
        try:
            with urllib.request.urlopen(url, timeout=10) as r, open(path + ".part", "wb") as f:
                f.write(r.read())
            os.replace(path + ".part", path)
            load_font(path, 40)
            return path
        except Exception:
            continue
    return None


def find_font():
    """폰트를 찾아요. 못 찾으면 한동안 다시 시도하지 않아서, 화면이 다시 그려질 때마다 멈추지 않아요"""
    if _FONT["path"] and os.path.exists(_FONT["path"]):
        return _FONT["path"]
    p = _system_font()
    if p:
        _FONT.update(path=p, source="서버에 설치된 폰트")
        return p
    if time.time() - _FONT["fail_at"] < FONT_RETRY_SEC:
        return None
    p = _download_font()
    if p:
        _FONT.update(path=p, source="자동 내려받음")
    else:
        _FONT["fail_at"] = time.time()
    return p


# ----------------------------------------------------------------------
# 파일 파싱
# ----------------------------------------------------------------------
ALIASES = {
    "hanja": ["한자", "hanja", "char", "character", "漢字"],
    "meaning": ["뜻", "훈", "meaning", "의미"],
    "sound": ["음", "독음", "sound", "reading", "소리"],
    "group": ["급수", "단원", "분류", "group", "level", "chapter", "category"],
}


def has_hanja(s):
    return any(
        "\u4e00" <= c <= "\u9fff" or "\u3400" <= c <= "\u4dbf" or "\uf900" <= c <= "\ufaff"
        for c in str(s)
    )


def decode_bytes(data):
    for enc in ("utf-8-sig", "cp949", "utf-16"):
        try:
            return data.decode(enc)
        except Exception:
            continue
    raise ValueError("파일 인코딩을 읽을 수 없어요. UTF-8로 저장해 주세요.")


_DELIM = re.compile(r"[,\t;|]")


def parse_text(text):
    """줄마다 '한자 뜻 음' (공백 구분) 또는 '한자,뜻,음,분류' (쉼표·탭 구분, 4번째 열은 분류)"""
    lines = [ln.strip().lstrip("\ufeff") for ln in text.splitlines() if ln.strip()]
    rows = []
    if lines and sum(1 for ln in lines if _DELIM.search(ln)) * 2 >= len(lines):
        cells = [[c.strip() for c in _DELIM.split(ln)] for ln in lines]
        width = max(len(c) for c in cells)
        for r in parse_table(pd.DataFrame([c + [""] * (width - len(c)) for c in cells])):
            h, m, s_ = r[0], r[1], r[2]
            if not s_.strip() and " " in m.strip():  # '天, 하늘 천' 처럼 뜻·음이 한 칸에 있는 경우
                m, s_ = m.strip().rsplit(None, 1)
            rows.append((h, m, s_) + tuple(r[3:]))
        return rows
    for ln in lines:
        toks = [t for t in re.split(r"[,\t;|\s]+", ln) if t]
        if len(toks) >= 3 and has_hanja(toks[0]):
            rows.append((toks[0], " ".join(toks[1:-1]), toks[-1]))
    return rows


def parse_table(df):
    df = df.fillna("").astype(str)
    if df.empty:
        return []
    first = [x.strip().lower() for x in df.iloc[0]]
    cols = {"hanja": 0, "meaning": 1, "sound": 2, "group": None}
    start = 0
    if df.shape[1] >= 4:  # 헤더가 없어도 4번째 열은 급수/단원/분류로 봄 (안내 문구와 같게)
        cols["group"] = 3
    if not any(has_hanja(x) for x in first):  # 첫 줄이 헤더
        cols["group"] = None
        start = 1
        for key, names in ALIASES.items():
            for i, h in enumerate(first):
                if h in [n.lower() for n in names]:
                    cols[key] = i
    rows = []
    for _, r in df.iloc[start:].iterrows():
        vals = list(r)

        def get(i):
            return vals[i].strip() if i < len(vals) else ""

        grp = get(cols["group"]) if cols["group"] is not None else ""
        rows.append((get(cols["hanja"]), get(cols["meaning"]), get(cols["sound"]), grp))
    return rows


def parse_file(f):
    name = f.name.lower()
    data = f.getvalue()
    rows = []
    if name.endswith((".xlsx", ".xls")):
        rows = parse_table(pd.read_excel(io.BytesIO(data), header=None, dtype=str))
    elif name.endswith(".csv"):
        try:
            df = pd.read_csv(
                io.StringIO(strip_comments(decode_bytes(data))), header=None, dtype=str, keep_default_na=False
            )
            rows = parse_table(df)
        except Exception:
            rows = parse_text(strip_comments(decode_bytes(data)))
    else:
        rows = parse_text(strip_comments(decode_bytes(data)))

    return build_words(rows)


def _split_raw(s):
    return [x.strip() for x in re.split(r"[,/;·]", str(s)) if x.strip()]


def build_words(rows):
    """줄 목록 -> 단어 목록. 같은 한자가 여러 번 나오면(예: 行 행/항) 뜻·음을 합쳐서 모두 정답으로 인정해요.
    반환: (단어들, 건너뛴 줄 수, 뜻·음을 합친 줄 수)"""
    words, index, dropped, merged = [], {}, 0, 0
    for row in rows:
        h, m, s_ = row[0], row[1], row[2]
        g = row[3] if len(row) > 3 else ""
        h = h.strip()
        if not has_hanja(h) or not s_.strip():
            dropped += 1
            continue
        if h in index:
            w = words[index[h]]
            changed = False
            if m.strip() and m.strip() not in _split_raw(w["meaning"]):
                w["meaning"] += "/" + m.strip()
                changed = True
            if s_.strip() not in _split_raw(w["sound"]):
                w["sound"] += "/" + s_.strip()
                changed = True
            if g.strip() and not w["group"]:
                w["group"] = str(g).strip()
            merged += int(changed)
            continue
        index[h] = len(words)
        words.append({"hanja": h, "meaning": m.strip(), "sound": s_.strip(), "group": str(g).strip()})
    return words, dropped, merged


def _loaded_msg(n, dropped, merged):
    msg = f"{n}개 한자를 불러왔어요."
    if dropped:
        msg += f" (뜻·음이 없어서 제외한 줄 {dropped}개)"
    if merged:
        msg += f" (같은 한자의 다른 뜻·음 {merged}건은 합쳤어요)"
    return msg


# ----------------------------------------------------------------------
# 뜻·음 채점
# ----------------------------------------------------------------------
def norm(s):
    return re.sub(r"[\s\-_.,/·]", "", str(s))


def split_alts(s):
    return [norm(x) for x in re.split(r"[,/;·]", str(s)) if norm(x)]


def check_reading(user, w):
    u = norm(user)
    sounds = split_alts(w["sound"])
    meanings = split_alts(w["meaning"]) + [norm(w["meaning"])]
    ok_sound, rest = False, u
    for sd in sounds:
        if u.endswith(sd):
            ok_sound, rest = True, u[: -len(sd)]
            break
    ok_meaning = rest in meanings and rest != ""
    if ok_sound and ok_meaning:
        return True, ""
    if ok_sound:
        return False, "음은 맞았어요. 뜻이 달라요."
    if ok_meaning:
        return False, "뜻은 맞았어요. 음이 달라요."
    return False, ""


# ----------------------------------------------------------------------
# 손글씨 채점 (골격 비교 + 위치·크기 보정 + 획 단위 검사 + 비슷한 글자 구분)
# ----------------------------------------------------------------------
GRID = 96  # 비교용 격자 크기
GPAD = 8  # 격자 가장자리 여백
INNER = GRID - 2 * GPAD  # 글자가 차지하는 영역 (80)

# 채점 설정의 기본값. 키 이름은 사이드바 위젯의 key 와 같음
GRADE_DEFAULTS = dict(
    set_thr=0.70,  # 통과 점수
    set_tol=4.0,  # 위치 허용 오차 (글자 크기의 %)
    set_align=1.5,  # 위치·크기·기울기 보정 강도 (0 = 끔)
    set_beta=1.0,  # 1보다 크면 '빠진 획'에, 작으면 '덧쓴 획'에 더 엄격
    set_chk=True,  # 획 단위 검사 (빠진 획 / 덧쓴 획)
    set_miss=12.0,  # 이 길이(글자 크기의 %) 이상 어긋난 획이 있으면 오답
    set_disc=True,  # 비슷한 다른 글자와 비교
    set_margin=0.03,  # 다른 글자가 이 점수 차 넘게 더 비슷하면 오답
    set_variants=True,  # 같은 한자의 다른 글꼴 모양(한국/중국/일본식)도 인정
)
GRADE_PRESETS = {
    "관대": dict(set_thr=0.64, set_tol=5.0, set_align=2.0, set_beta=1.0, set_chk=True,
                 set_miss=15.0, set_disc=True, set_margin=0.05, set_variants=True),
    "보통": dict(GRADE_DEFAULTS),
    "엄격": dict(set_thr=0.80, set_tol=3.0, set_align=1.0, set_beta=1.2, set_chk=True,
                 set_miss=8.0, set_disc=True, set_margin=0.01, set_variants=False),
}
N_CANDIDATES = 6  # 비슷한 글자 비교 때 정밀하게 따져 볼 후보 수
STROKE_BAD = 2.0  # 허용 오차의 몇 배 넘게 떨어진 부분을 '어긋남'으로 볼지


def grade_params():
    s = st.session_state
    return {k: s.get(k, v) for k, v in GRADE_DEFAULTS.items()}


@st.cache_resource(show_spinner=False)
def _store():
    """폰트·정답 글자 분석 결과를 앱이 다시 실행돼도 유지하는 캐시"""
    return {}


def _font(path, size, index=None):
    key = ("font", path, size, index)
    store = _store()
    f = store.get(key)
    if f is None:
        f = load_font(path, size) if index is None else ImageFont.truetype(path, size, index=index)
        store[key] = f
    return f


def font_variants(path, use_variants=True):
    """같은 한자라도 폰트 지역판(한국/중국/일본)마다 모양이 조금 달라서, 가능하면 여러 모양을 모두 정답으로 인정"""
    if not path:
        return []
    if not use_variants or not path.lower().endswith(".ttc"):
        return [(path, None)]
    out = []
    for idx in (1, 3, 0):  # Noto CJK: 1=KR, 3=TC, 0=JP
        try:
            _font(path, 40, idx)
            out.append((path, idx))
        except Exception:
            continue
    return out or [(path, None)]


def render_char(ch, font_path, size=256, index=None):
    font = _font(font_path, int(size * 0.8), index)
    img = Image.new("L", (size, size), 0)
    ImageDraw.Draw(img).text((size / 2, size / 2), ch, font=font, fill=255, anchor="mm")
    return img


GRID_LINE = (232, 196, 192)  # 연한 붉은 점선 (채점 때 획으로 인식되지 않는 밝기)


def grid_image(size=320):
    """손글씨 칸에 깔아 두는 십자 점선(田字格). 글씨 균형을 잡는 데 도움이 돼요"""
    base = Image.new("RGB", (size, size), (255, 255, 255))
    d = ImageDraw.Draw(base)
    c = size // 2
    for t in range(0, size, 14):
        d.line([(c, t), (c, min(t + 7, size))], fill=GRID_LINE, width=1)
        d.line([(t, c), (min(t + 7, size), c)], fill=GRID_LINE, width=1)
    return base


def guide_image(ch, font_path, size=320):
    mask = np.array(render_char(ch, font_path, size)) > 127
    base = np.array(grid_image(size))
    base[mask] = (215, 215, 215)  # 연한 회색 (채점 때 획으로 인식되지 않는 밝기)
    return Image.fromarray(base)


def thin(mask):
    """Zhang-Suen 세선화: 굵은 글자/가는 손글씨를 같은 굵기(1픽셀 골격)로 맞춤"""
    a = np.pad(mask.astype(np.uint8), 1)
    changed = True
    while changed:
        changed = False
        for step in (0, 1):
            p = a
            P2, P3, P4, P5 = p[:-2, 1:-1], p[:-2, 2:], p[1:-1, 2:], p[2:, 2:]
            P6, P7, P8, P9 = p[2:, 1:-1], p[2:, :-2], p[1:-1, :-2], p[:-2, :-2]
            c = p[1:-1, 1:-1]
            B = P2 + P3 + P4 + P5 + P6 + P7 + P8 + P9
            seq = [P2, P3, P4, P5, P6, P7, P8, P9, P2]
            A = sum(((seq[i] == 0) & (seq[i + 1] == 1)).astype(np.uint8) for i in range(8))
            if step == 0:
                m = (P2 * P4 * P6 == 0) & (P4 * P6 * P8 == 0)
            else:
                m = (P2 * P4 * P8 == 0) & (P2 * P6 * P8 == 0)
            d = (c == 1) & (B >= 2) & (B <= 6) & (A == 1) & m
            if d.any():
                a[1:-1, 1:-1][d] = 0
                changed = True
    return a[1:-1, 1:-1].astype(bool)


def _thick_mask(sub):
    """글자 영역(bool)을 비율을 유지한 채 GRID 크기 정사각형에 맞춰 넣음 (겹쳐 보기·후보 고르기용)"""
    h, w = sub.shape
    side = max(w, h)
    sq = Image.new("L", (side, side), 0)
    sq.paste(Image.fromarray((sub * 255).astype("uint8")), ((side - w) // 2, (side - h) // 2))
    sq = sq.resize((INNER, INNER), Image.BILINEAR)
    canvas = Image.new("L", (GRID, GRID), 0)
    canvas.paste(sq, (GPAD, GPAD))
    return np.array(canvas) > 100


def _dist_map(pts):
    """격자의 모든 칸에서 가장 가까운 골격 점까지의 거리"""
    g = np.arange(GRID, dtype=np.float32)
    gx, gy = np.meshgrid(g, g)
    gx, gy = gx.reshape(-1, 1), gy.reshape(-1, 1)
    out = np.empty(GRID * GRID, np.float32)
    for i in range(0, GRID * GRID, 1024):
        d2 = (gx[i : i + 1024] - pts[:, 0][None, :]) ** 2 + (gy[i : i + 1024] - pts[:, 1][None, :]) ** 2
        out[i : i + 1024] = np.sqrt(d2.min(1))
    return out.reshape(GRID, GRID)


def prep_shape(mask):
    """글씨(bool 이미지) -> 크기를 맞춘 골격 점들 + 거리 맵. 글자가 없으면 None"""
    ys, xs = np.nonzero(mask)
    if len(ys) < 20:
        return None
    x0, x1, y0, y1 = xs.min(), xs.max() + 1, ys.min(), ys.max() + 1
    w, h = int(x1 - x0), int(y1 - y0)
    sub = mask[y0:y1, x0:x1]
    sy, sx = np.nonzero(thin(sub))
    if len(sy) < 3:
        return None
    scale = INNER / max(w, h)
    pts = np.stack([(sx + 0.5 - w / 2) * scale + GRID / 2, (sy + 0.5 - h / 2) * scale + GRID / 2], 1)
    pts = pts.astype(np.float32)
    if len(pts) > 900:
        pts = pts[:: -(-len(pts) // 900)]
    return {"pts": pts, "thick": _thick_mask(sub), "dt": _dist_map(pts)}


TPL_CACHE_MAX = 500  # 서버 메모리 보호용: 보관하는 정답 글자 분석 결과 수


def get_templates(ch, font_path, use_variants=True):
    """정답 글자의 골격(글꼴 변형별). 한 번 계산하면 캐시에 보관"""
    key = ("tpl", font_path, ch, bool(use_variants))
    store = _store()
    if key in store:
        return store[key]
    out, seen = [], set()
    for path, idx in font_variants(font_path, use_variants):
        try:
            arr = np.array(render_char(ch, path, 192, idx)) > 127
        except Exception:
            continue
        sig = hash(arr.tobytes())
        if sig in seen:
            continue
        seen.add(sig)
        shp = prep_shape(arr)
        if shp is not None:
            out.append(shp)
    store[key] = out
    tpl_keys = [k for k in list(store) if k[0] == "tpl"]
    if len(tpl_keys) > TPL_CACHE_MAX:  # 오래된 절반을 비움 (다시 필요하면 자동으로 다시 계산)
        for k in tpl_keys[: len(tpl_keys) // 2]:
            store.pop(k, None)
    return out


def _lowres(thick):
    im = Image.fromarray((thick * 255).astype("uint8")).resize((24, 24), Image.BOX)
    v = np.asarray(im, np.float32).ravel()
    v = v - v.mean()
    return v / (np.linalg.norm(v) + 1e-6)


def _vocab_feat(ch, font_path):
    key = ("vf", font_path, ch)
    store = _store()
    if key not in store:
        v = None
        try:
            arr = np.array(render_char(ch, font_path, 128)) > 127
            ys, xs = np.nonzero(arr)
            if len(ys) >= 20:
                v = _lowres(_thick_mask(arr[ys.min() : ys.max() + 1, xs.min() : xs.max() + 1]))
        except Exception:
            pass
        store[key] = v
    return store[key]


# --- 위치·크기·기울기 보정 -------------------------------------------------
# 변환 T = (dx, dy, log sx, log sy, 각도(rad)) : 내 글씨 점을 이동·확대·회전해서 정답에 맞춰 봄
def _steps(level):
    a, k, r = 1.6 * level, 0.06 * level, np.deg2rad(2.0 * level)
    return a, a, k, k, r


def _grid_around(center, steps, frac=1.0):
    ax = [np.array([-1.0, 0.0, 1.0]) * s * frac + c for c, s in zip(center, steps)]
    mesh = np.meshgrid(*ax, indexing="ij")
    return np.stack([m.ravel() for m in mesh], 1).astype(np.float32)


def _apply(pts, T):
    c = GRID / 2
    x = (pts[:, 0] - c)[None, :] * np.exp(T[:, 2])[:, None]
    y = (pts[:, 1] - c)[None, :] * np.exp(T[:, 3])[:, None]
    co, si = np.cos(T[:, 4])[:, None], np.sin(T[:, 4])[:, None]
    return co * x - si * y + T[:, 0:1] + c, si * x + co * y + T[:, 1:2] + c


def _apply_inv(pts, T):
    c = GRID / 2
    x = (pts[:, 0] - c)[None, :] - T[:, 0:1]
    y = (pts[:, 1] - c)[None, :] - T[:, 1:2]
    co, si = np.cos(T[:, 4])[:, None], np.sin(T[:, 4])[:, None]
    return (co * x + si * y) / np.exp(T[:, 2])[:, None] + c, (-si * x + co * y) / np.exp(T[:, 3])[:, None] + c


def _lookup(dm, x, y):
    ix = np.clip(np.rint(x), 0, GRID - 1).astype(np.int32)
    iy = np.clip(np.rint(y), 0, GRID - 1).astype(np.int32)
    return dm[iy, ix]


def _eval(u, t, T, tol, beta):
    """모든 변환 후보에 대해 점수 계산. 내 글씨 점 -> 정답까지(정밀도), 정답 점 -> 내 글씨까지(재현율)"""
    ux, uy = _apply(u["pts"], T)
    du = _lookup(t["dt"], ux, uy)
    tx, ty = _apply_inv(t["pts"], T)
    dt = _lookup(u["dt"], tx, ty)
    p = np.clip(2.0 - du / tol, 0, 1).mean(1)  # tol 이내 1점, 2*tol 에서 0점까지 부드럽게 감소
    r = np.clip(2.0 - dt / tol, 0, 1).mean(1)
    b2 = beta * beta
    f = (1 + b2) * p * r / (b2 * p + r + 1e-9)
    shift = np.hypot(T[:, 0], T[:, 1]) / INNER
    pen = 0.30 * shift + 0.10 * (np.abs(T[:, 2]) + np.abs(T[:, 3])) + 0.10 * np.abs(T[:, 4])
    return f - pen, p, r, du, dt


def _pick(u, t, T, tol, beta):
    sc, p, r, du, dt = _eval(u, t, T, tol, beta)
    i = int(np.argmax(sc))
    return dict(score=float(sc[i]), p=float(p[i]), r=float(r[i]), T=T[i], du=du[i], dt=dt[i])


def _search(u, t, tol, beta, level, refine=False):
    """거친 탐색(+선택적으로 세밀 탐색)으로 가장 잘 맞는 위치·크기·기울기를 찾음"""
    if level <= 0:
        return _pick(u, t, np.zeros((1, 5), np.float32), tol, beta)
    st_ = _steps(level)
    best = _pick(u, t, _grid_around((0, 0, 0, 0, 0), st_), tol, beta)
    if refine:
        for frac in (0.5, 0.25):
            cand = _pick(u, t, _grid_around(best["T"], st_, frac), tol, beta)
            if cand["score"] > best["score"]:
                best = cand
    return best


def _max_cluster_extent(pts, bad, link=2.4):
    """어긋난 점들 중 서로 이어진 가장 큰 덩어리의 크기(대각선 길이)"""
    q = pts[bad]
    n = len(q)
    if n < 2:
        return 0.0
    adj = ((q[:, None, :] - q[None, :, :]) ** 2).sum(-1) <= link * link
    seen = np.zeros(n, bool)
    best = 0.0
    for i in range(n):
        if seen[i]:
            continue
        seen[i] = True
        stack, comp = [i], [i]
        while stack:
            j = stack.pop()
            for k in np.nonzero(adj[j] & ~seen)[0]:
                seen[k] = True
                stack.append(k)
                comp.append(k)
        c = q[comp]
        best = max(best, float(np.hypot(*(c.max(0) - c.min(0)))))
    return best


def _aligned_user_mask(thick, T):
    """겹쳐 보기용: 찾아낸 보정값을 내 글씨에 적용"""
    ix, iy = _apply_inv(np.array([[0, 0], [1, 0], [0, 1]], np.float32), T[None, :])
    (x0, y0), (x1, y1), (x2, y2) = [(float(ix[0, i]), float(iy[0, i])) for i in range(3)]
    coef = (x1 - x0, x2 - x0, x0, y1 - y0, y2 - y0, y0)
    img = Image.fromarray((thick * 255).astype("uint8")).transform((GRID, GRID), Image.AFFINE, coef, Image.BILINEAR)
    return np.array(img) > 100


def overlay(u, t, scale=4):
    h, w = u.shape
    rgb = np.full((h, w, 3), 255, np.uint8)
    rgb[t] = (255, 190, 190)
    rgb[u] = (30, 60, 220)
    rgb[u & t] = (120, 40, 160)
    return Image.fromarray(rgb).resize((w * scale, h * scale), Image.NEAREST)


def grade_writing(image_data, ch, font_path, P, candidates=()):
    """손글씨 채점. 결과 dict (글씨를 못 알아보면 None)

    1) 내 글씨와 정답 글자의 골격을 위치·크기·기울기를 보정하며 비교  -> 점수
    2) 빠진 획 / 덧쓴 획(이어진 큰 어긋남)이 있는지 검사
    3) 같은 단어장의 비슷한 글자와 비교해 그쪽이 더 비슷하면 오답
    """
    gray = image_data[:, :, :3].astype(float).mean(axis=2)
    u = prep_shape(gray < 110)
    if u is None:
        return None
    tpls = get_templates(ch, font_path, P["set_variants"])
    if not tpls:
        return None
    tol = max(0.5, P["set_tol"] * INNER / 100.0)
    beta, level = float(P["set_beta"]), float(P["set_align"])

    # 정답 글자: 글꼴 변형 중 가장 비슷한 것을 골라 세밀하게 보정
    coarse = [_search(u, t, tol, beta, level) for t in tpls]
    ib = int(np.argmax([c["score"] for c in coarse]))
    tpl, s_coarse = tpls[ib], coarse[ib]["score"]
    best = _search(u, tpl, tol, beta, level, refine=True)
    score = best["score"]

    # 비슷한 다른 글자 (모양이 가장 닮은 몇 개만 정밀 비교)
    other = None
    if P["set_disc"] and candidates:
        uv = _lowres(u["thick"])
        sims = []
        for c in dict.fromkeys(candidates):
            if c == ch:
                continue
            v = _vocab_feat(c, font_path)
            if v is not None:
                sims.append((float(uv @ v), c))
        sims.sort(reverse=True)
        for _, c in sims[:N_CANDIDATES]:
            ts = get_templates(c, font_path, P["set_variants"])
            if not ts:
                continue
            sc = max(_search(u, t, tol, beta, level)["score"] for t in ts)
            if other is None or sc > other[1]:
                other = (c, sc)

    # 획 단위 검사
    miss = extra = 0.0
    if P["set_chk"]:
        bad = STROKE_BAD * tol
        miss = _max_cluster_extent(tpl["pts"], best["dt"] > bad) / INNER * 100
        ux, uy = _apply(u["pts"], best["T"][None, :])
        extra = _max_cluster_extent(np.stack([ux[0], uy[0]], 1), best["du"] > bad) / INNER * 100

    ok, reason, detail = score >= P["set_thr"], "", ""
    if not ok:
        reason = "score"
    elif P["set_chk"] and miss >= P["set_miss"]:
        ok, reason, detail = False, "stroke", "정답에 있는 획이 빠졌거나 많이 달라요."
    elif P["set_chk"] and extra >= P["set_miss"]:
        ok, reason, detail = False, "stroke", "정답에 없는 획이 더 들어갔어요."
    elif other is not None and other[1] - s_coarse > P["set_margin"]:
        ok, reason = False, "disc"
        detail = f"'{other[0]}' 글자와 더 비슷하게 써졌어요. (정답 {s_coarse:.2f} < {other[0]} {other[1]:.2f})"
    if not ok and not detail and other is not None and other[1] - s_coarse > P["set_margin"]:
        detail = f"'{other[0]}' 글자와 더 비슷해요."

    ua = _aligned_user_mask(u["thick"], best["T"])
    return dict(
        ok=ok, score=score, cmp=s_coarse, p=best["p"], r=best["r"], reason=reason, detail=detail,
        miss=miss, extra=extra, other=other, overlay=overlay(ua, tpl["thick"]),
        shift=(float(best["T"][0]), float(best["T"][1])),
    )


# ----------------------------------------------------------------------
# 상태 관리
# ----------------------------------------------------------------------
INTERVALS = [0, 1, 3, 7, 14, 30]  # 복습 상자(0~5)별로 다음 복습까지 며칠 뒤인지
BOX_LABELS = ["바로 복습", "1일 뒤", "3일 뒤", "7일 뒤", "14일 뒤", "30일 뒤"]
TYPE_LABELS = {
    "read": "뜻·음 쓰기",
    "write": "한자 쓰기",
    "choice_ko": "객관식 (한자→뜻·음)",
    "choice_hj": "객관식 (뜻·음→한자)",
}
SETTING_DEFAULTS = dict(
    set_type="섞어서", set_source="전체 단어", set_count=10, set_shuffle=True, set_groups=[],
    set_goal=20, set_stroke=8, set_canvas_px=180,
    card_src="전체 단어", card_dir="한자 → 뜻·음", card_shuffle=True, card_count=20,
    # 시험 플랜
    set_exam_date="", set_target=90, plan_groups=[],
    # 사자성어
    idm_view="목록", idm_src="전체", idm_type="섞어서", idm_count=10,
)
KEEP_KEYS = list(SETTING_DEFAULTS)  # 화면에 없는 동안에도 값이 사라지지 않게 붙잡아 둘 설정들


def esc(x):
    return htmllib.escape(str(x))


KST = timezone(timedelta(hours=9))  # 서버는 세계 표준시(UTC)라서, 한국 시간으로 '오늘'을 계산해요


def today_date():
    return datetime.now(KST).date()


def today_str():
    return today_date().isoformat()


def init_state():
    defaults = dict(
        words=[], file_sig=None, quiz=[], pos=0, results=[], feedback=None,
        wrong={}, stats={}, nonce=0, retry_mode=False, hint=False, flash="",
        progress_sig=None, calib=[], set_preset="보통", srs={}, history={}, cards=None, sheet=None,
        restored=False, store_ok=None,
        notes={}, idioms=[], idiom_sig=None, idiom_stats={}, idiom_wrong={},
        iquiz=[], ipos=0, ires=[], inonce=0, iretry=False,
        ed_ver=0, notice=None, hw_clear=0, up_gen=0, wipe_id=0,
    )
    defaults.update(GRADE_DEFAULTS)
    defaults.update(SETTING_DEFAULTS)
    for k, v in defaults.items():
        st.session_state.setdefault(k, copy.deepcopy(v))
    # 위젯이 화면에 없는 실행에서도 설정값이 지워지지 않게 다시 저장
    for k in KEEP_KEYS:
        st.session_state[k] = st.session_state[k]


# --- 복습 주기 (간격 반복) ---------------------------------------------------
def srs_update(key, ok):
    """맞히면 상자가 한 칸 올라가 복습 간격이 늘고, 틀리면 0번 상자로 돌아와 바로 복습 대상이 돼요"""
    s = st.session_state
    e = s.srs.get(key, {"box": 0, "due": today_str(), "seen": 0})
    box = min(e.get("box", 0) + 1, len(INTERVALS) - 1) if ok else 0
    due = (today_date() + timedelta(days=INTERVALS[box])).isoformat()
    s.srs[key] = {"box": box, "due": due, "seen": e.get("seen", 0) + 1}


def due_pool(words):
    """오늘 복습할 차례인 한자: 복습 날짜가 지난 것(약한 것 먼저) → 아직 안 풀어 본 것"""
    s = st.session_state
    today = today_str()
    review, new = [], []
    for w in words:
        e = s.srs.get(w["hanja"])
        if e is None:
            new.append(w)
        elif e.get("due", today) <= today:
            review.append((e.get("box", 0), e.get("due", ""), w))
    review.sort(key=lambda x: (x[0], x[1]))
    random.shuffle(new)
    return [w for _, _, w in review] + new


def due_counts(words):
    s = st.session_state
    today = today_str()
    new = sum(1 for w in words if w["hanja"] not in s.srs)
    review = sum(1 for w in words if w["hanja"] in s.srs and s.srs[w["hanja"]].get("due", today) <= today)
    return new, review


# --- 날짜별 기록 ---------------------------------------------------------------
def hist_add(day, right=0, wrong=0):
    h = st.session_state.history.setdefault(day, {"right": 0, "wrong": 0})
    h["right"] += right
    h["wrong"] += wrong


def study_streak():
    """오늘(아직 안 풀었으면 어제)부터 거꾸로, 하루도 안 빠지고 푼 날이 며칠인지"""
    h = st.session_state.history

    def done(d):
        v = h.get(d.isoformat(), {})
        return v.get("right", 0) + v.get("wrong", 0) > 0

    d, n = today_date(), 0
    if not done(d):
        d -= timedelta(days=1)
    while done(d) and n < 3650:
        n += 1
        d -= timedelta(days=1)
    return n


def hist_today():
    h = st.session_state.history.get(today_str(), {})
    return h.get("right", 0) + h.get("wrong", 0)


# --- 문제 만들기 -----------------------------------------------------------------
def reset_round(quiz, retry=False):
    s = st.session_state
    s.quiz, s.pos, s.results = quiz, 0, []
    s.feedback, s.hint, s.flash = None, False, ""
    s.retry_mode = retry
    s.nonce += 1


def _label(w):
    return f"{w['meaning']} {w['sound']}".strip()


def make_choices(w, ty, n=4):
    """보기 만들기: 같은 급수/단원의 한자를 먼저 오답 보기로 써서 헷갈리게 만들어요"""
    s = st.session_state
    others = [x for x in s.words if x["hanja"] != w["hanja"]]
    if ty == "choice_hj":  # 뜻·음이 같은 다른 한자는 정답 시비가 생기니 제외
        others = [x for x in others if _label(x) != _label(w)]
    grp = w.get("group")
    same = [x for x in others if grp and x.get("group") == grp]
    rest = [x for x in others if not (grp and x.get("group") == grp)]
    random.shuffle(same)
    random.shuffle(rest)
    key = (lambda x: _label(x)) if ty == "choice_ko" else (lambda x: x["hanja"])
    seen, pick = {key(w)}, []
    for x in same + rest:
        if key(x) in seen:
            continue
        seen.add(key(x))
        pick.append(x)
        if len(pick) == n - 1:
            break
    opts = pick + [w]
    random.shuffle(opts)
    return [{"hanja": x["hanja"], "label": _label(x)} for x in opts]


def make_item(w, ty):
    item = {"w": w, "type": ty}
    if ty in ("choice_ko", "choice_hj"):
        item["opts"] = make_choices(w, ty)
    return item


def start_quiz(pool_override=None, count_override=None):
    s = st.session_state
    source = "오늘 복습할 것" if pool_override is not None else s.get("set_source", "전체 단어")
    kind = s.get("set_type", "섞어서")
    groups = s.get("set_groups", [])
    if pool_override is not None:  # 시험 플랜: 오늘 복습 + 새 한자를 묶어서 시작
        pool = list(pool_override)
    elif source == "오답 노트만":
        pool = list(s.wrong.values())
    else:
        pool = [w for w in s.words if not groups or w.get("group") in groups]
        if source == "별표 한자만":
            pool = [w for w in pool if is_star(w["hanja"])]
        if source == "오늘 복습할 것":
            pool = due_pool(pool)
    if not pool:
        if source == "오늘 복습할 것":
            s.flash = "오늘 복습할 한자가 없어요. 내일 다시 만나요! 🎉"
        elif source == "별표 한자만":
            s.flash = "별표를 준 한자가 없어요. 단어장 탭에서 ★를 눌러 주세요."
        else:
            s.flash = "문제로 만들 단어가 없어요. 파일을 올리거나 범위·오답 노트를 확인해 주세요."
        return

    many = len(s.words) >= 4  # 객관식은 보기를 만들려면 한자가 4개 이상 필요
    note = ""
    if kind in ("한자 쓰기", "섞어서") and not HAS_CANVAS:
        if kind == "한자 쓰기":
            s.flash = "⚠️ 한자 쓰기 기능(streamlit-drawable-canvas)을 불러오지 못해서 시작할 수 없어요. 왼쪽 메뉴의 '진단 정보'를 확인해 주세요."
            return
        note = "⚠️ 한자 쓰기 기능을 불러오지 못해 손글씨 문제는 빠졌어요. (왼쪽 메뉴 '진단 정보' 확인)"
    if kind == "객관식" and not many:
        s.flash = "객관식은 단어장에 한자가 4개 이상 있어야 만들 수 있어요."
        return

    def pick_type(w):
        can_write = HAS_CANVAS and len(w["hanja"]) == 1
        if kind == "뜻·음 쓰기":
            return "read"
        if kind == "한자 쓰기":
            return "write" if can_write else None
        if kind == "객관식":
            return random.choice(["choice_ko", "choice_hj"])
        opts = ["read"] + (["write"] if can_write else []) + (["choice_ko", "choice_hj"] if many else [])
        return random.choice(opts)

    shuffle = s.get("set_shuffle", True)
    if shuffle and source != "오늘 복습할 것":
        random.shuffle(pool)
    count = int(count_override or s.get("set_count", 10))
    items = []
    for w in pool:
        ty = pick_type(w)
        if ty:
            items.append(make_item(w, ty))
        if len(items) >= count:
            break
    if not items:
        s.flash = "선택한 유형으로 만들 수 있는 문제가 없어요. (한자 쓰기는 한 글자 한자만 가능)"
        return
    if shuffle and source == "오늘 복습할 것":
        random.shuffle(items)
    reset_round(items, retry=(source == "오답 노트만"))
    s.flash = note


def record(w, qtype, ok, ans):
    s = st.session_state
    key = w["hanja"]
    day = today_str()
    prev = s.srs.get(key)
    entry = {
        "w": w, "type": qtype, "correct": ok, "answer": ans, "added_wrong": False,
        "day": day, "srs_prev": dict(prev) if prev else None,
    }
    st_ = s.stats.setdefault(key, {"meaning": w["meaning"], "sound": w["sound"], "right": 0, "wrong": 0})
    if ok:
        st_["right"] += 1
        if s.retry_mode and key in s.wrong:
            del s.wrong[key]
    else:
        st_["wrong"] += 1
        e = s.wrong.setdefault(
            key, {"hanja": key, "meaning": w["meaning"], "sound": w["sound"], "wrong_count": 0, "last": ""}
        )
        e["wrong_count"] += 1
        e["last"] = datetime.now(KST).strftime("%Y-%m-%d %H:%M")
        entry["added_wrong"] = True
    hist_add(day, right=int(ok), wrong=int(not ok))
    srs_update(key, ok)
    s.results.append(entry)


def override_correct():
    """채점이 억울할 때 '맞게 썼어요'로 정정"""
    s = st.session_state
    last = s.results[-1]
    if last["correct"]:
        return
    key = last["w"]["hanja"]
    if last["type"] == "write" and (s.feedback or {}).get("score") is not None:
        s.calib.append({"score": s.feedback["score"], "why": s.feedback.get("reason", "")})
    last["correct"] = True
    s.stats[key]["wrong"] -= 1
    s.stats[key]["right"] += 1
    if last["added_wrong"] and key in s.wrong:
        s.wrong[key]["wrong_count"] -= 1
        if s.wrong[key]["wrong_count"] <= 0:
            del s.wrong[key]
    hist_add(last.get("day", today_str()), right=1, wrong=-1)
    prev = last.get("srs_prev")
    if prev:
        s.srs[key] = dict(prev)
    else:
        s.srs.pop(key, None)
    srs_update(key, True)
    s.feedback["ok"] = True
    s.feedback["detail"] = "정정되었어요."


def next_question():
    s = st.session_state
    s.pos += 1
    s.feedback, s.hint = None, False
    s.nonce += 1


def reset_quiz():
    """진행 중인 퀴즈만 초기화 (오답 노트·통계는 유지)"""
    s = st.session_state
    s.quiz, s.pos, s.results = [], 0, []
    s.feedback, s.hint = None, False
    s.nonce += 1
    s.flash = "퀴즈를 마쳤어요. 설정을 바꿔서 다시 시작할 수 있어요."


def factory_reset():
    """앱을 처음 상태로: 단어장·기록·메모·설정·진행 중인 퀴즈, 그리고 이 브라우저의 자동 저장까지 전부 지워요"""
    s = st.session_state
    gen, wipe, store_ok = s.get("up_gen", 0) + 1, s.get("wipe_id", 0) + 1, s.get("store_ok")
    for k in list(s.keys()):
        del s[k]
    s.up_gen = gen  # 파일 올리는 칸도 새것으로 (올려 둔 파일이 다시 읽히지 않게)
    s.wipe_id = wipe  # 브라우저 저장소를 비우라는 신호
    s.restored, s.store_ok = True, store_ok  # 지운 옛 기록을 다시 불러오지 않게
    s.flash = "앱을 처음 상태로 되돌렸어요."


def reset_all():
    """퀴즈 + 오답 노트 + 통계 + 복습 기록 전부 초기화 (불러온 단어는 유지)"""
    s = st.session_state
    s.wrong, s.stats, s.srs, s.history = {}, {}, {}, {}
    s.idiom_stats, s.idiom_wrong = {}, {}
    s.cards = None
    reset_idiom_quiz()
    reset_quiz()
    s.confirm_reset = False
    s.flash = "퀴즈, 오답 노트, 통계, 복습 기록을 모두 초기화했어요. (단어장·메모는 그대로예요)"


# ----------------------------------------------------------------------
# 새 기능 공통: 상한 · 검증 · 폴더 내장 세트 · 메모/별표 · 단어장 편집
# ----------------------------------------------------------------------
MAX_NOTE = 200  # 메모 한 개의 최대 글자 수
MAX_NOTES = 3000  # 저장하는 메모·별표 개수 (브라우저 저장 용량을 지키려는 상한)
MAX_IDIOMS = 2000  # 사자성어 개수
IDIOM_MIN, IDIOM_MAX = 3, 8  # 사자성어 한자 글자 수
MAX_SET_BYTES = 2_000_000  # 폴더 CSV 한 개의 최대 크기

IDIOM_LABELS = {
    "i_read": "사자성어 → 독음 쓰기",
    "i_ch_snd": "독음 → 사자성어 고르기",
    "i_ch_ko": "사자성어 → 뜻 고르기",
    "i_ch_hj": "뜻 → 사자성어 고르기",
    "i_blank": "빈칸 채우기 (고르기)",
    "i_hw": "빈칸 한 글자 손글씨",
}
IDIOM_KINDS = ["섞어서", "독음 쓰기", "독음 → 사자성어", "뜻 고르기", "빈칸 고르기", "빈칸 손글씨"]
IDIOM_KIND_TYPES = {
    "독음 쓰기": ["i_read"],
    "독음 → 사자성어": ["i_ch_snd"],
    "뜻 고르기": ["i_ch_ko", "i_ch_hj"],
    "빈칸 고르기": ["i_blank"],
    "빈칸 손글씨": ["i_hw"],
}
IDIOM_SRC = ["전체", "오답 노트만"]
SOURCES = ["전체 단어", "오늘 복습할 것", "오답 노트만", "별표 한자만"]


def _dict(x):
    return x if isinstance(x, dict) else {}


def _flt(x, default=0.0, lo=0.0, hi=1e9):
    try:
        v = float(x)
    except Exception:
        return default
    if not math.isfinite(v):
        return default
    return max(lo, min(hi, v))


def _cell(x):
    """표 칸 값 -> 문자열 (비어 있으면 '')"""
    try:
        if x is None or pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x).strip()


def _all_hanja(x):
    return bool(x) and all(has_hanja(c) for c in x)


def fmt_pts(x):
    return f"{float(x):.1f}".rstrip("0").rstrip(".")


def fmt_secs(n):
    n = int(n)
    return f"{n // 60}분 {n % 60}초" if n >= 60 else f"{n}초"


def now_kst_str():
    return datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")


def strip_comments(text):
    """'#'으로 시작하는 줄은 메모로 보고 무시해요 (예시 파일의 '검수 필요' 안내용)"""
    return "\n".join(ln for ln in text.splitlines() if not ln.lstrip("\ufeff \t").startswith("#"))


# --- 검증 ---------------------------------------------------------------------
def word_problem(h, m, sd, g=""):
    """한자 한 개 입력의 문제점 (없으면 '')"""
    if not h:
        return "한자가 비어 있어요"
    if not has_hanja(h):
        return "한자가 아닌 글자예요"
    if len(h) > 12:
        return "한자가 너무 길어요 (12자 이하)"
    if not sd:
        return "음이 비어 있어요"
    if len(m) > 100 or len(sd) > 60 or len(g) > 30:
        return "뜻(100자)·음(60자)·분류(30자)가 너무 길어요"
    return ""


def idiom_problem(h, sd, mn):
    if not h:
        return "사자성어(한자)가 비어 있어요"
    if not _all_hanja(h):
        return "한자가 아닌 글자가 섞여 있어요"
    if not (IDIOM_MIN <= len(h) <= IDIOM_MAX):
        return f"한자는 {IDIOM_MIN}~{IDIOM_MAX}자여야 해요 (지금 {len(h)}자)"
    if not sd:
        return "독음이 비어 있어요"
    if not mn:
        return "뜻이 비어 있어요"
    if len(sd) > 60 or len(mn) > 300:
        return "독음(60자)이나 뜻(300자)이 너무 길어요"
    return ""


def _clean_idiom(it):
    if not isinstance(it, dict):
        return None
    h = re.sub(r"\s+", "", str(it.get("hanja", "")))
    sd, mn = str(it.get("sound", "")).strip(), str(it.get("meaning", "")).strip()
    if idiom_problem(h, sd, mn):
        return None
    return {"hanja": h, "sound": sd, "meaning": mn, "group": str(it.get("group", "")).strip()[:30]}


def _clean_idioms(lst):
    out, seen = [], set()
    for it in lst if isinstance(lst, list) else []:
        c = _clean_idiom(it)
        if c and c["hanja"] not in seen:
            seen.add(c["hanja"])
            out.append(c)
            if len(out) >= MAX_IDIOMS:
                break
    return out


def _clean_notes(d):
    out = {}
    for k, v in _dict(d).items():
        if len(out) >= MAX_NOTES:
            break
        if not isinstance(v, dict) or not has_hanja(str(k)):
            continue
        note = str(v.get("note", "") or "").strip()[:MAX_NOTE]
        star = v.get("star") is True
        if note or star:
            out[str(k)[:30]] = {"note": note, "star": star}
    return out


def _clean_idiom_key(k):
    k = str(k)
    return k if (IDIOM_MIN <= len(k) <= IDIOM_MAX and _all_hanja(k)) else None


# --- 같은 항목 합치기 ----------------------------------------------------------
def _same_text(a, b):
    """b의 뜻(음)이 a에 이미 들어 있으면 같은 걸로 봐요 ('/'로 이은 여러 답도 따로 비교)"""
    if not norm(b):
        return True
    have = set(split_alts(a)) | {norm(a)}
    return all(x in have for x in (split_alts(b) or [norm(b)]))


def merge_items(base, incoming, fields):
    """같은 한자는 한 번만. 뜻·음이 서로 다르면 먼저 있던 것을 유지하고 conflicts에 모아요.
    반환: (합친 목록, 새로 들어온 수, 같아서 건너뛴 수, 뜻·음이 달라 유지한 한자들)"""
    out = [dict(x) for x in base]
    idx = {x["hanja"]: i for i, x in enumerate(out)}
    added = dup = 0
    conflicts = []
    for it in incoming:
        h = it["hanja"]
        if h in idx:
            cur = out[idx[h]]
            if all(_same_text(cur.get(f, ""), it.get(f, "")) for f in fields):
                dup += 1
            else:
                conflicts.append(h)
            if not cur.get("group") and it.get("group"):
                cur["group"] = it["group"]
            continue
        idx[h] = len(out)
        out.append(dict(it))
        added += 1
    return out, added, dup, conflicts


# --- 폴더(wordsets/, idioms/)의 내장 세트 --------------------------------------
def _app_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()


WORDSET_DIR = os.path.join(_app_dir(), "wordsets")
IDIOM_DIR = os.path.join(_app_dir(), "idioms")
SET_KINDS = {
    "word": dict(dir=WORDSET_DIR, folder="wordsets", label="단어장", obj="단어장을", fields=("meaning", "sound")),
    "idiom": dict(dir=IDIOM_DIR, folder="idioms", label="사자성어", obj="사자성어를", fields=("meaning", "sound")),
}


def set_title(fname):
    return os.path.splitext(fname)[0].replace("_", " ").strip() or fname


def list_sets(folder):
    """폴더의 CSV 목록 -> [(세트 이름, 경로)]. 폴더가 없거나 비어 있으면 빈 목록"""
    try:
        names = sorted(n for n in os.listdir(folder)
                       if n.lower().endswith(".csv") and not n.startswith((".", "~")))
    except Exception:
        return []
    out, seen = [], set()
    for n in names[:100]:
        p = os.path.join(folder, n)
        if not os.path.isfile(p):
            continue
        t = set_title(n)
        if t in seen:  # 이름이 겹치면 파일 이름 그대로
            t = n
        seen.add(t)
        out.append((t, p))
    return out


@st.cache_data(show_spinner=False, max_entries=64)
def _read_set_cached(path, mtime, size, kind):
    with open(path, "rb") as f:
        data = f.read(MAX_SET_BYTES + 1)
    if len(data) > MAX_SET_BYTES:
        raise ValueError("파일이 너무 커요 (2MB 이하만 읽어요)")
    text = strip_comments(decode_bytes(data))
    parse = parse_table if kind == "word" else parse_idiom_table
    try:
        df = pd.read_csv(io.StringIO(text), header=None, dtype=str, keep_default_na=False)
        return parse(df)
    except Exception:
        return parse_text(text) if kind == "word" else parse_idiom_text(text)


def read_set_rows(path, kind):
    try:
        info = os.stat(path)
        return _read_set_cached(path, info.st_mtime, info.st_size, kind), ""
    except Exception as e:
        return [], f"{type(e).__name__}: {e}"


def clear_notice():
    st.session_state.notice = None


def show_notice(kind, tag):
    s = st.session_state
    n = s.get("notice")
    if isinstance(n, dict) and n.get("kind") == kind and n.get("msgs"):
        for m in n["msgs"]:
            st.warning(m)
        st.button("확인했어요", key=f"notice_ok_{tag}", on_click=clear_notice)


def load_builtin(kind, tag):
    """고른 내장 세트를 불러와요 (바꾸기 / 합치기). 버튼 콜백으로 실행돼요"""
    s = st.session_state
    cfg = SET_KINDS[kind]
    sets = dict(list_sets(cfg["dir"]))
    picked = [n for n in s.get(f"ws_pick_{tag}", []) if n in sets]
    if not picked:
        s.flash = f"불러올 {cfg['obj']} 먼저 골라 주세요."
        return
    notes, conflicts, incoming = [], [], []
    for name in picked:
        rows, err = read_set_rows(sets[name], kind)
        if err:
            notes.append(f"'{name}'을(를) 읽지 못했어요: {err}")
            continue
        if kind == "word":
            items, dropped, _m = build_words(rows)
            if dropped:
                notes.append(f"'{name}': 뜻·음이 없어서 제외한 줄 {dropped}개")
        else:
            items, probs, _d = build_idioms(rows)
            if probs:
                notes.append(f"'{name}': 제외한 줄 {len(probs)}개 (예: {probs[0]})")
        incoming, _a, _d2, conf = merge_items(incoming, items, cfg["fields"])
        conflicts += conf
    if not incoming:
        s.notice = {"kind": kind, "msgs": notes or ["읽을 수 있는 항목이 없었어요."]}
        s.flash = "읽을 수 있는 항목이 없었어요."
        return
    mode = s.get(f"ws_mode_{tag}", "바꾸기")
    base = (s.words if kind == "word" else s.idioms) if mode != "바꾸기" else []
    merged, added, dup, conf2 = merge_items(base, incoming, cfg["fields"])
    conflicts += conf2
    if kind == "word":
        s.words = merged
        reset_views()
    else:
        if len(merged) > MAX_IDIOMS:
            notes.append(f"사자성어는 최대 {MAX_IDIOMS}개까지만 담아요. 넘는 {len(merged) - MAX_IDIOMS}개는 뺐어요.")
        s.idioms = merged[:MAX_IDIOMS]
        reset_idiom_quiz()
    if conflicts:
        notes.append(f"같은 한자인데 뜻·음이 달라서 '먼저 있던 것'을 유지한 항목이 {len(conflicts)}개 있어요: "
                     + ", ".join(conflicts[:10]) + (" …" if len(conflicts) > 10 else ""))
    s.notice = {"kind": kind, "msgs": notes} if notes else None
    msg = f"{cfg['label']} {len(picked)}개 세트를 불러왔어요. (새 항목 {added}개 · 겹쳐서 한 번만 담은 것 {dup}개)"
    if conflicts:
        msg += " ⚠️ 뜻·음이 달라 유지한 항목이 있어요. 안내를 확인해 주세요."
    s.flash = msg


def set_picker(kind, tag, show=True):
    """내장 세트 고르기 (처음 화면 / 단어장 탭 / 사자성어 관리에서 같이 써요)"""
    s = st.session_state
    cfg = SET_KINDS[kind]
    if show:
        show_notice(kind, tag)
    sets = list_sets(cfg["dir"])
    if not sets:
        st.caption(f"📂 내장 {cfg['label']}이 아직 없어요. 앱 폴더의 {cfg['folder']}/ 폴더에 CSV 파일을 넣으면 여기에 목록이 나와요.")
        return
    names = [n for n, _ in sets]
    pk, mk = f"ws_pick_{tag}", f"ws_mode_{tag}"
    s[pk] = [x for x in s.get(pk, []) if x in names]
    merge_label = "지금 단어장에 합치기" if kind == "word" else "지금 사자성어에 합치기"
    modes = ["바꾸기", merge_label]
    if s.get(mk) not in modes:
        s[mk] = merge_label if (s.words if kind == "word" else s.idioms) else "바꾸기"
    with st.form(f"ws_form_{tag}", border=False):
        st.multiselect(f"내장 {cfg['label']} 고르기 (여러 개 가능)", names, key=pk, placeholder="눌러서 고르기")
        st.radio("불러오는 방법", modes, key=mk, horizontal=True,
                 help="같은 한자는 한 번만 담아요. 뜻·음이 서로 다르면 먼저 있던 것을 유지하고 알려 드려요.")
        st.form_submit_button(f"📂 {cfg['label']} 불러오기", type="primary", on_click=load_builtin, args=(kind, tag), width="stretch")
    st.caption(f"{cfg['folder']}/ 폴더의 CSV {len(sets)}개 · 파일 이름이 세트 이름이에요. 열: "
               + ("한자, 뜻, 음, 분류(선택)" if kind == "word" else "사자성어, 독음, 뜻, 분류(선택)"))


# --- 사자성어 읽기 ---------------------------------------------------------------
IDIOM_ALIASES = {
    "hanja": ["사자성어", "성어", "한자", "hanja", "idiom"],
    "sound": ["독음", "음", "읽기", "sound", "reading"],
    "meaning": ["뜻", "의미", "해석", "meaning"],
    "group": ["분류", "단원", "급수", "group", "category", "chapter"],
}


def parse_idiom_table(df):
    df = df.fillna("").astype(str)
    if df.empty:
        return []
    first = [x.strip().lower() for x in df.iloc[0]]
    cols = {"hanja": 0, "sound": 1, "meaning": 2, "group": 3 if df.shape[1] >= 4 else None}
    start = 0
    if not any(has_hanja(x) for x in first):  # 첫 줄이 헤더
        start, cols["group"] = 1, None
        for key, names in IDIOM_ALIASES.items():
            for i, h in enumerate(first):
                if h in [n.lower() for n in names]:
                    cols[key] = i
    rows = []
    for _, r in df.iloc[start:].iterrows():
        vals = list(r)

        def get(i):
            return vals[i].strip() if i is not None and i < len(vals) else ""

        rows.append((get(cols["hanja"]), get(cols["sound"]), get(cols["meaning"]), get(cols["group"])))
    return rows


def parse_idiom_text(text, group=""):
    """한 줄에 '一石二鳥 일석이조 한 가지 일로 두 가지 이익을 얻음' (공백 구분).
    탭·'|'·';'로 나눠도 되고, '한자,독음,뜻'처럼 쉼표도 돼요. '#'으로 시작하는 줄은 무시해요."""
    rows = []
    for ln in text.splitlines():
        ln = ln.strip().lstrip("\ufeff")
        if not ln or ln.startswith("#"):
            continue
        if re.search(r"[\t|;]", ln):
            parts = [p.strip() for p in re.split(r"[\t|;]", ln, maxsplit=3)]
        elif "," in ln and _all_hanja(ln.split(",", 1)[0].strip()):
            parts = [p.strip() for p in ln.split(",", 2)]
        else:
            parts = ln.split(None, 2)
        parts += [""] * (4 - len(parts))
        rows.append((parts[0], parts[1], parts[2], parts[3] or group))
    return rows


def parse_idiom_file(f):
    name, data = f.name.lower(), f.getvalue()
    if name.endswith((".xlsx", ".xls")):
        return parse_idiom_table(pd.read_excel(io.BytesIO(data), header=None, dtype=str))
    text = strip_comments(decode_bytes(data))
    if name.endswith(".csv"):
        try:
            return parse_idiom_table(pd.read_csv(io.StringIO(text), header=None, dtype=str, keep_default_na=False))
        except Exception:
            pass
    return parse_idiom_text(text)


def build_idioms(rows):
    """줄 목록 -> (사자성어 목록, 문제 있는 줄 설명들, 같은 사자성어 중복 수)"""
    out, idx, problems, dup = [], {}, [], 0
    for n, row in enumerate(rows, 1):
        h = re.sub(r"\s+", "", row[0])
        sd, mn = row[1].strip(), row[2].strip()
        g = (row[3] if len(row) > 3 else "").strip()
        if not (h or sd or mn):
            continue
        why = idiom_problem(h, sd, mn)
        if why:
            problems.append(f"{n}번째 줄 ({h or '한자 없음'}): {why}")
            continue
        if h in idx:
            dup += 1
            if g and not out[idx[h]]["group"]:
                out[idx[h]]["group"] = g[:30]
            continue
        idx[h] = len(out)
        out.append({"hanja": h, "sound": sd, "meaning": mn, "group": g[:30]})
    return out, problems, dup


# --- 메모 · 별표 -----------------------------------------------------------------
def is_star(h):
    return bool(st.session_state.notes.get(h, {}).get("star"))


def note_of(h):
    return st.session_state.notes.get(h, {}).get("note", "")


def _put_note(h, note, star):
    s = st.session_state
    note = str(note).strip()[:MAX_NOTE]
    if note or star:
        if h not in s.notes and len(s.notes) >= MAX_NOTES:
            return False
        s.notes[h] = {"note": note, "star": bool(star)}
    else:
        s.notes.pop(h, None)
    return True


def toggle_star(h, on):
    _put_note(h, note_of(h), on)


def bulk_star(on):
    s = st.session_state
    n = 0
    for h in s.get("nt_bulk", []):
        n += int(_put_note(h, note_of(h), on))
    s.nt_bulk = []
    s.flash = f"{n}개 한자에 별표를 {'줬어요' if on else '뺐어요'}."


def save_note(h, text):
    ok = _put_note(h, text, is_star(h))
    st.toast("메모를 저장했어요." if ok else f"메모·별표는 최대 {MAX_NOTES}개까지만 저장해요.")


def show_memo(h):
    n = note_of(h)
    if n:
        html(f"<div class='memo'>📝 내 메모<br>{esc(n)}</div>")


# --- 단어장 편집 -------------------------------------------------------------------
WORD_COLS = ["한자", "뜻", "음", "분류"]


def words_df(words):
    return pd.DataFrame(
        [{"한자": w["hanja"], "뜻": w["meaning"], "음": w["sound"], "분류": w.get("group", "")} for w in words],
        columns=WORD_COLS,
    )


def words_csv(words):
    return words_df(words).to_csv(index=False).encode("utf-8-sig")


def words_from_df(df):
    """편집 표 -> (단어 목록, 오류 목록). 오류가 하나라도 있으면 저장하지 않아요"""
    words, errors, seen = [], [], {}
    for i, (_, r) in enumerate(df.iterrows(), 1):
        h = re.sub(r"\s+", "", _cell(r.get("한자")))
        m, sd, g = _cell(r.get("뜻")), _cell(r.get("음")), _cell(r.get("분류"))
        if not (h or m or sd or g):
            continue  # 완전히 빈 줄은 무시
        why = word_problem(h, m, sd, g)
        if why:
            errors.append(f"{i}번째 줄 ({h or '한자 없음'}): {why}")
            continue
        if h in seen:
            errors.append(f"{i}번째 줄 ({h}): {seen[h]}번째 줄과 같은 한자예요. 다른 음은 '행/항'처럼 / 로 이어서 한 줄에 적어 주세요")
            continue
        seen[h] = i
        words.append({"hanja": h, "meaning": m, "sound": sd, "group": g})
    return words, errors


def reset_views():
    """단어장이 바뀔 때: 진행 중이던 퀴즈·카드·시험지를 안전하게 비워요 (기록은 그대로)"""
    s = st.session_state
    s.quiz, s.pos, s.results, s.feedback, s.cards, s.sheet = [], 0, [], None, None, None
    s.hint = False
    s.nonce += 1


def reset_idiom_quiz():
    s = st.session_state
    s.iquiz, s.ipos, s.ires = [], 0, []
    s.inonce += 1


def commit_words(new_words):
    """편집 결과를 단어장에 반영해요. 뜻·음을 고치면 오답 노트·통계에 복사돼 있는 값도 같이 바꿔요.
    (한자 글자를 바꾸면 '삭제 + 새 단어'가 돼요. 삭제한 한자의 기록은 지우지 않고 남겨 둬요)"""
    s = st.session_state
    old = {w["hanja"]: w for w in s.words}
    new_keys = {w["hanja"] for w in new_words}
    changed = 0
    for w in new_words:
        o = old.get(w["hanja"])
        if o and (o["meaning"] != w["meaning"] or o["sound"] != w["sound"]):
            changed += 1
            for book in (s.wrong, s.stats):
                e = book.get(w["hanja"])
                if e:
                    e["meaning"], e["sound"] = w["meaning"], w["sound"]
    added = sum(1 for w in new_words if w["hanja"] not in old)
    removed = sum(1 for h in old if h not in new_keys)
    s.words = new_words
    reset_views()
    s.ed_ver += 1
    s.flash = (f"단어장을 저장했어요. (추가 {added} · 수정 {changed} · 삭제 {removed}) "
               "진행 중이던 퀴즈·카드는 안전하게 비웠어요. 새로 시작해 주세요.")


def delete_words(keys, with_records=False):
    """고른 한자를 단어장에서 지워요. with_records=True 면 그 한자의 오답·통계·복습·메모 기록도 같이 지워요.
    단어장이 비게 되면 시작 화면으로 돌아가요."""
    s = st.session_state
    keys = set(keys)
    before = len(s.words)
    s.words = [w for w in s.words if w["hanja"] not in keys]
    n = before - len(s.words)
    if with_records:
        for book in (s.wrong, s.stats, s.srs, s.notes):
            for k in keys:
                book.pop(k, None)
    reset_views()
    s.ed_ver += 1
    if not s.words:  # 비었으면 파일 올리는 칸도 새것으로 (같은 파일을 다시 올릴 수 있게)
        s.up_gen += 1
        s.file_sig = None
    s.flash = f"한자 {n}개를 삭제했어요." + (" 단어장이 비었어요." if not s.words else "")


def delete_selected_cb():
    s = st.session_state
    keys = list(s.get("del_pick", []))
    if keys:
        delete_words(keys, bool(s.get("del_rec", False)))
    s.del_pick = []
    s.del_rec = False


def delete_all_cb():
    s = st.session_state
    delete_words([w["hanja"] for w in s.words], bool(s.get("del_rec", False)))
    s.del_pick, s.del_rec, s.del_all_ok = [], False, False


def add_words_from_text(text):
    """붙여넣은 글의 한자를 지금 단어장 뒤에 이어 붙여요. 이미 있는 한자는 새 뜻·음만 '/'로 덧붙여요.
    반환: 사용자에게 보여줄 안내 문구, 성공 여부"""
    s = st.session_state
    rows, bad = [], 0
    for r in parse_text(strip_comments(text or "")):
        g = (r[3] if len(r) > 3 else "").strip()
        if word_problem(r[0].strip(), r[1].strip(), r[2].strip(), g):
            bad += 1
        else:
            rows.append(r)
    if not rows:
        return "읽을 수 있는 줄이 없어요. 한 줄에 '한자 뜻 음' 모양인지 확인해 주세요.", False
    old = [(w["hanja"], w["meaning"], w["sound"], w.get("group", "")) for w in s.words]
    words, _, _ = build_words(old + rows)
    if words == s.words:
        return "새로 추가할 내용이 없어요. (이미 단어장에 있는 한자예요)", False
    commit_words(words)  # 추가·수정 개수 안내와 퀴즈 정리는 여기서 해요
    if bad:
        s.flash += f" (형식이 맞지 않아 건너뛴 줄 {bad}개)"
    return s.flash, True


def paste_add_cb(text_key):
    s = st.session_state
    msg, ok = add_words_from_text(s.get(text_key, ""))
    if ok:
        s[text_key] = ""  # 성공하면 입력칸을 비워요
    else:
        s.flash = msg


# ----------------------------------------------------------------------
# 자동 저장 (이 기기의 브라우저에 보관) + 불러온 기록 검증
# ----------------------------------------------------------------------
STORE_KEY = "hanja-practice-v1"
STORE_HTML = """<!doctype html><html><head><meta charset="utf-8"></head><body style="margin:0"><script>
(function () {
  var KEY = "__KEY__";
  function send(type, data) {
    var m = { isStreamlitMessage: true, type: type };
    for (var k in data) { m[k] = data[k]; }
    window.parent.postMessage(m, "*");
  }
  var loaded = false, last = null, lastClear = 0;
  function canStore() {
    try {
      window.localStorage.setItem("__hanja_test", "1");
      window.localStorage.removeItem("__hanja_test");
      return true;
    } catch (e) { return false; }
  }
  function readStored() {
    try { return window.localStorage.getItem(KEY) || ""; } catch (e) { return ""; }
  }
  window.addEventListener("message", function (ev) {
    var d = ev.data;
    if (!d || d.type !== "streamlit:render") { return; }
    var a = d.args || {};
    if (typeof a.clear === "number" && a.clear > lastClear) {  // 앱 완전 초기화: 저장돼 있던 기록 삭제
      lastClear = a.clear;
      try { window.localStorage.removeItem(KEY); } catch (e) {}
      last = null;
    }
    if (!loaded) {
      loaded = true;
      send("streamlit:setComponentValue", {
        value: JSON.stringify({ ok: canStore(), data: readStored() }), dataType: "json"
      });
    }
    if (typeof a.save === "string" && a.save !== "" && a.save !== last) {
      try { window.localStorage.setItem(KEY, a.save); last = a.save; } catch (e) {}
    }
  });
  send("streamlit:componentReady", { apiVersion: 1 });
  send("streamlit:setFrameHeight", { height: 0 });
})();
</script></body></html>"""

PERSIST_KEYS = [
    "set_goal", "set_stroke", "set_canvas_px", "set_type", "set_source", "set_count", "set_shuffle", "set_preset",
    "card_src", "card_dir", "card_shuffle", "card_count",
    "set_exam_date", "set_target", "plan_groups", "idm_src", "idm_type", "idm_count",
] + list(GRADE_DEFAULTS)
# 저장값 범위 (엉뚱한 값이 들어 있어도 위젯이 깨지지 않게 이 안으로 맞춰요)
SETTING_RANGES = {"set_target": (0, 100), "idm_count": (5, 50)}
ALLOWED = {
    "set_type": ["섞어서", "뜻·음 쓰기", "한자 쓰기", "객관식"],
    "set_source": SOURCES,
    "card_src": SOURCES,
    "idm_src": IDIOM_SRC,
    "idm_type": IDIOM_KINDS,
    "card_dir": ["한자 → 뜻·음", "뜻·음 → 한자"],
    "set_preset": ["관대", "보통", "엄격", "직접 설정"],
}


def _int(x, default=0, lo=0, hi=10**9):
    try:
        return max(lo, min(hi, int(x)))
    except Exception:
        return default


def _iso(x):
    try:
        return date.fromisoformat(str(x)).isoformat()
    except Exception:
        return None


def _clean_word(w):
    if not isinstance(w, dict):
        return None
    h, sd = str(w.get("hanja", "")).strip(), str(w.get("sound", "")).strip()
    if not has_hanja(h) or not sd:
        return None
    return {"hanja": h, "meaning": str(w.get("meaning", "")).strip(), "sound": sd,
            "group": str(w.get("group", "")).strip()}


def sanitize_progress(d):
    """파일/브라우저에서 읽은 기록을 검사해서, 엉뚱한 값이 섞여 있어도 앱이 깨지지 않게 걸러 줘요"""
    out = {"wrong": {}, "stats": {}, "srs": {}, "history": {}, "notes": {},
           "idiom_stats": {}, "idiom_wrong": {}, "plan": {}}
    if not isinstance(d, dict):
        return out
    for k, v in _dict(d.get("wrong")).items():
        if isinstance(v, dict) and has_hanja(str(k)):
            out["wrong"][str(k)] = {
                "hanja": str(k), "meaning": str(v.get("meaning", "")), "sound": str(v.get("sound", "")),
                "wrong_count": _int(v.get("wrong_count"), 1, 1), "last": str(v.get("last", "")),
            }
    for k, v in _dict(d.get("stats")).items():
        if isinstance(v, dict) and has_hanja(str(k)):
            out["stats"][str(k)] = {
                "meaning": str(v.get("meaning", "")), "sound": str(v.get("sound", "")),
                "right": _int(v.get("right")), "wrong": _int(v.get("wrong")),
            }
    for k, v in _dict(d.get("srs")).items():
        due = _iso(v.get("due")) if isinstance(v, dict) else None
        if due and has_hanja(str(k)):
            out["srs"][str(k)] = {"box": _int(v.get("box"), 0, 0, len(INTERVALS) - 1), "due": due,
                                  "seen": _int(v.get("seen"))}
    for day, v in _dict(d.get("history")).items():
        if isinstance(v, dict) and _iso(day):
            out["history"][_iso(day)] = {"right": _int(v.get("right")), "wrong": _int(v.get("wrong"))}
    # 새 항목 (옛 저장값에는 없으니 없으면 빈 값)
    out["notes"] = _clean_notes(d.get("notes"))
    for k, v in _dict(d.get("idiom_stats")).items():
        if isinstance(v, dict) and _clean_idiom_key(k):
            out["idiom_stats"][str(k)] = {"sound": str(v.get("sound", "")), "meaning": str(v.get("meaning", "")),
                                          "right": _int(v.get("right")), "wrong": _int(v.get("wrong"))}
    for k, v in _dict(d.get("idiom_wrong")).items():
        if isinstance(v, dict) and _clean_idiom_key(k) and str(v.get("sound", "")).strip() and str(v.get("meaning", "")).strip():
            out["idiom_wrong"][str(k)] = {
                "hanja": str(k), "sound": str(v.get("sound", "")), "meaning": str(v.get("meaning", "")),
                "wrong_count": _int(v.get("wrong_count"), 1, 1), "last": str(v.get("last", ""))}
    p = _dict(d.get("plan"))
    if "exam_date" in p:
        out["plan"]["exam_date"] = _iso(p["exam_date"]) or ""
    if "target" in p:
        out["plan"]["target"] = _int(p["target"], 90, 0, 100)
    if isinstance(p.get("groups"), list):
        out["plan"]["groups"] = [x[:60] for x in p["groups"] if isinstance(x, str)][:100]
    return out


def merge_progress(clean):
    s = st.session_state
    s.wrong.update(clean["wrong"])
    s.stats.update(clean["stats"])
    s.srs.update(clean["srs"])
    for day, v in clean["history"].items():
        cur = s.history.setdefault(day, {"right": 0, "wrong": 0})
        cur["right"] = max(cur["right"], v["right"])
        cur["wrong"] = max(cur["wrong"], v["wrong"])
    s.notes.update(clean.get("notes", {}))
    s.idiom_stats.update(clean.get("idiom_stats", {}))
    s.idiom_wrong.update(clean.get("idiom_wrong", {}))
    plan = clean.get("plan") or {}
    if "exam_date" in plan:
        s.set_exam_date = plan["exam_date"]
    if "target" in plan:
        s.set_target = plan["target"]
    if "groups" in plan:
        s.plan_groups = plan["groups"]


def progress_json():
    s = st.session_state
    return json.dumps(
        {"wrong": s.wrong, "stats": s.stats, "srs": s.srs, "history": s.history,
         "notes": s.notes, "idiom_stats": s.idiom_stats, "idiom_wrong": s.idiom_wrong,
         "plan": {"exam_date": s.get("set_exam_date", ""), "target": s.get("set_target", 90),
                  "groups": s.get("plan_groups", [])}},
        ensure_ascii=False, indent=2,
    ).encode("utf-8")


def build_payload():
    s = st.session_state
    data = {"v": 2, "words": s.words, "wrong": s.wrong, "stats": s.stats, "srs": s.srs,
            "history": s.history, "notes": s.notes, "idiom_stats": s.idiom_stats,
            "idiom_wrong": s.idiom_wrong,
            "settings": {k: s[k] for k in PERSIST_KEYS if k in s}}
    return json.dumps(data, ensure_ascii=False, separators=(",", ":"), default=str)


def apply_payload(raw):
    """브라우저에 저장돼 있던 단어장·기록·설정을 되살려요. 성공하면 True"""
    s = st.session_state
    try:
        data = json.loads(raw)
    except Exception:
        return False
    if not isinstance(data, dict):
        return False
    if not s.words:
        seen, words = set(), []
        for w in data.get("words") or []:
            c = _clean_word(w)
            if c and c["hanja"] not in seen:
                seen.add(c["hanja"])
                words.append(c)
        s.words = words
    try:
        merge_progress(sanitize_progress(data))
    except Exception:  # 저장값이 이상해도 앱이 멈추지 않게
        pass
    defaults = dict(GRADE_DEFAULTS)
    defaults.update(SETTING_DEFAULTS)
    defaults["set_preset"] = "보통"
    for k, v in (data.get("settings") or {}).items():
        if k in PERSIST_KEYS and k in defaults and type(v) is type(defaults[k]):
            if k in ALLOWED and v not in ALLOWED[k]:
                continue
            if k == "set_canvas_px":  # 엉뚱한 값이 저장돼 있어도 슬라이더가 깨지지 않게 범위 안으로
                v = _int(v, 180, 120, 360)
            if k == "set_exam_date":
                v = _iso(v) or ""
            if k == "plan_groups":
                v = [x[:60] for x in v if isinstance(x, str)][:100]
            if k in SETTING_RANGES:
                lo, hi = SETTING_RANGES[k]
                v = _int(v, defaults[k], lo, hi) if isinstance(defaults[k], int) else _flt(v, defaults[k], lo, hi)
                if k == "idm_count":
                    v = max(5, min(50, int(round(v / 5.0)) * 5))
            s[k] = v
    return bool(s.words or s.wrong or s.stats or s.srs or s.history or s.notes
                or s.idiom_stats or s.idiom_wrong)


@st.cache_resource(show_spinner=False)
def _store_component():
    d = os.path.join(tempfile.gettempdir(), "hanja_browser_store")
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, "index.html"), "w", encoding="utf-8") as f:
        f.write(STORE_HTML.replace("__KEY__", STORE_KEY))
    return components.declare_component("hanja_browser_store", path=d)


def sync_browser_store():
    """매 실행마다 현재 상태를 브라우저에 저장하고, 새로 열었을 땐 저장돼 있던 기록을 되살려요.
    실패해도 앱은 평소처럼 동작해요(저장만 안 될 뿐)."""
    s = st.session_state
    save = ""
    if s.restored and (s.words or s.wrong or s.stats or s.srs or s.history or s.notes
                       or s.idiom_stats or s.idiom_wrong):
        save = build_payload()  # 불러오기 전에는 저장하지 않아서, 빈 상태가 기존 기록을 덮어쓰지 않아요
    try:
        with st.container(key="store_box"):
            got = _store_component()(save=save, clear=int(s.get("wipe_id", 0)), key="browser_store", default=None)
    except Exception:
        s.restored, s.store_ok = True, False
        return
    if got is None or s.restored:
        return
    s.restored = True
    try:
        info = json.loads(got) if isinstance(got, str) else dict(got)
    except Exception:
        info = {}
    s.store_ok = bool(info.get("ok"))
    raw = info.get("data", "")
    if raw and apply_payload(raw):
        s.flash = f"이전 학습 기록을 불러왔어요. (한자 {len(s.words)}개)" if s.words else "이전 학습 기록을 불러왔어요."


def repair_state():
    """화면이 꼬였을 때: 진행 중이던 퀴즈·카드만 비워요 (단어장·기록은 그대로)"""
    s = st.session_state
    s.quiz, s.pos, s.results, s.feedback, s.cards, s.sheet = [], 0, [], None, None, None
    s.hint = False
    s.nonce += 1
    s.iquiz, s.ipos, s.ires = [], 0, []
    s.flash = "화면을 정리했어요."



# --- 플래시카드 ---------------------------------------------------------------
def new_deck(pool):
    deck = [{"hanja": w["hanja"], "meaning": w["meaning"], "sound": w["sound"]} for w in pool]
    return {"deck": deck, "i": 0, "flip": False, "known": [], "unsure": []}


def start_cards():
    s = st.session_state
    src = s.get("card_src", "전체 단어")
    if src == "오답 노트만":
        pool = list(s.wrong.values())
    elif src == "오늘 복습할 것":
        pool = due_pool(s.words)
    elif src == "별표 한자만":
        pool = [w for w in s.words if is_star(w["hanja"])]
    else:
        pool = list(s.words)
    if not pool:
        s.flash = ("오늘 복습할 한자가 없어요. 🎉" if src == "오늘 복습할 것"
                   else "별표를 준 한자가 없어요. 단어장 탭에서 ★를 눌러 주세요." if src == "별표 한자만"
                   else "카드로 만들 한자가 없어요.")
        return
    if s.get("card_shuffle", True) and src != "오늘 복습할 것":
        random.shuffle(pool)
    s.cards = new_deck(pool[: int(s.get("card_count", 20))])
    s.flash = ""


def card_flip():
    st.session_state.cards["flip"] = True


def card_mark(known):
    s = st.session_state
    c = s.cards
    cur = c["deck"][c["i"]]
    (c["known"] if known else c["unsure"]).append(cur)
    srs_update(cur["hanja"], known)
    c["i"] += 1
    c["flip"] = False


def card_retry_unsure():
    c = st.session_state.cards
    st.session_state.cards = new_deck(c["unsure"])


def card_close():
    st.session_state.cards = None


# ----------------------------------------------------------------------
# 디자인 (CSS)
# ----------------------------------------------------------------------
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@500;700&family=Noto+Sans+KR:wght@400;500;700&display=swap');
:root{
  --ink:#1E2438; --indigo:#2B3A67; --jade:#2F7D6D; --seal:#B8372C;
  --mist:#E7EBF2; --line:#D3D9E4; --soft:#5B6479; --grid:rgba(184,55,44,.38);
  --serif:'Noto Serif KR','Noto Serif CJK KR','Batang','Songti SC',serif;
}
.stApp, .stApp .stMarkdown, .stApp label, .stApp button, .stApp input, .stApp textarea{
  font-family:'Noto Sans KR','Apple SD Gothic Neo','Malgun Gothic',system-ui,sans-serif;
}
.stApp h1,.stApp h2,.stApp h3{font-family:var(--serif); letter-spacing:-.01em;}
.block-container{padding-top:4.2rem; padding-bottom:5rem; max-width:1000px;}

/* 머리글 */
.app-head{display:flex; align-items:center; gap:14px; margin:0 0 .6rem;}
.app-head .seal{width:52px; height:52px; border-radius:10px; background:var(--seal); color:#fff;
  display:grid; place-items:center; font:700 31px/1 var(--serif); transform:rotate(-3deg); flex:none;}
.app-head .title{font:700 1.9rem/1.15 var(--serif); color:inherit;}
.app-head .sub{color:inherit; opacity:.7; font-size:.95rem; margin-top:2px;}
.chips{display:flex; flex-wrap:wrap; gap:8px; margin:.2rem 0 .7rem;}
.chip{display:inline-flex; align-items:center; gap:6px; padding:.3rem .85rem; border-radius:999px;
  background:var(--mist); color:var(--ink); font-weight:500; font-size:.95rem;}
.chip.hot{background:#F7E3E0; color:var(--seal);}
.chip.good{background:#DCEEE9; color:var(--jade);}

/* 터치 친화 */
.stButton>button,.stDownloadButton>button{min-height:3.1rem; border-radius:12px; font-weight:500; font-size:1.02rem;}
.stTabs [data-baseweb="tab-list"]{gap:.2rem;}
.stTabs [data-baseweb="tab"]{height:3.2rem; padding:0 1.05rem; font-size:1.02rem;}
.stRadio label p, .stCheckbox label p, .stToggle label p{font-size:1.02rem;}
.stRadio [role="radiogroup"]{gap:.5rem;}
.stRadio [role="radiogroup"] label{padding:.5rem 1rem; margin:0; border:1.5px solid var(--line); border-radius:999px; background:#fff;}
.stRadio [role="radiogroup"] label:has(input:checked){border-color:var(--indigo); background:#E9EDF8;}
.app-head.compact{margin-bottom:.2rem;}
.app-head.compact .seal{width:38px; height:38px; font-size:23px; border-radius:8px;}
.app-head.compact .title{font-size:1.35rem;}

/* 한자 카드: 십자 점선 칸 */
.hj-card{width:min(240px,70vw); aspect-ratio:1/1; margin:.2rem auto .8rem; display:flex; align-items:center;
  justify-content:center; font-family:var(--serif); font-weight:500; color:var(--ink); line-height:1;
  background:
    repeating-linear-gradient(to bottom,var(--grid) 0 6px,transparent 6px 12px) center/1px 100% no-repeat,
    repeating-linear-gradient(to right,var(--grid) 0 6px,transparent 6px 12px) center/100% 1px no-repeat,
    #fff;
  border:2px solid var(--indigo); border-radius:6px;}
.hj-card.wide{aspect-ratio:auto; width:fit-content; min-width:min(240px,70vw); max-width:100%;
  padding:1.2rem 1.2rem; background:#fff;}
.mean-card{max-width:520px; margin:.3rem auto .8rem; padding:1.1rem 1rem; text-align:center; background:#fff;
  border:2px solid var(--indigo); border-radius:6px; font-family:var(--serif); font-size:2rem; line-height:1.35; color:var(--ink);}
.mean-card b{display:block; font-size:2.3rem; margin-top:.15rem; color:var(--indigo);}
.ask{text-align:center; color:inherit; opacity:.75; margin:.1rem 0 .9rem; font-size:1.05rem;}

/* 결과: 도장 */
.fb-row{display:flex; align-items:center; justify-content:center; gap:1.4rem; flex-wrap:wrap; margin:.6rem 0 1rem;}
.verdict{display:inline-block; padding:.3rem 1.1rem; border:3px solid; border-radius:10px;
  font:700 1.7rem/1.2 var(--serif); transform:rotate(-4deg); letter-spacing:.05em;}
.verdict.ok{color:var(--jade); border-color:var(--jade);}
.verdict.no{color:var(--seal); border-color:var(--seal);}
.reveal{display:flex; align-items:center; gap:1rem; padding:.8rem 1.3rem; border:1px solid var(--line);
  border-radius:10px; background:#fff;}
.reveal .rh{font:500 3.2rem/1 var(--serif); color:var(--ink);}
.reveal .rm{font-size:1.15rem; color:var(--ink);}
.reveal .rm b{font:700 1.5rem var(--serif);}

/* 객관식 버튼 */
.st-key-choices_ko button,.st-key-choices_hj button{min-height:4.4rem; background:#fff; border:1.5px solid var(--line);}
.st-key-choices_ko button p{font-size:1.3rem;}
.st-key-choices_hj button p{font-family:var(--serif); font-size:2.6rem; line-height:1.1;}

/* 단어장 */
.wgrid{display:grid; grid-template-columns:repeat(auto-fill,minmax(128px,1fr)); gap:12px; margin-top:.6rem;}
.wcard{background:#fff; border:1px solid var(--line); border-radius:10px; padding:.7rem .4rem .6rem; text-align:center;}
.wcard .h{font:500 2.6rem/1.1 var(--serif); color:var(--ink);}
.wcard .m{font-size:.95rem; color:var(--soft); margin-top:.25rem; min-height:1.3em;}
.wcard .m b{color:var(--ink);}
.wcard details summary{cursor:pointer; color:var(--indigo); font-size:.9rem; list-style:none; padding:.25rem 0;}
.wcard details summary::-webkit-details-marker{display:none;}
.wcard .g{font-size:.75rem; color:var(--soft);}
.wcard mark{background:#FFE9A8; color:inherit; border-radius:3px; padding:0 2px;}

/* 결과 요약 */
.score-box{text-align:center; padding:1.2rem; border:2px solid var(--indigo); border-radius:6px; background:#fff; margin-bottom:1rem;}
.score-box .big{font:700 3.2rem/1.1 var(--serif); color:var(--indigo);}
.score-box .lbl{color:var(--soft);}

/* 어두운 화면 설정(다크 모드)에서도 글씨가 보이게: 흰 배경 위의 글씨는 색을 직접 지정 */
.stRadio [role="radiogroup"] label, .stRadio [role="radiogroup"] label p{color:var(--ink);}
.st-key-choices_ko button, .st-key-choices_hj button, .st-key-choices_ko button p, .st-key-choices_hj button p{color:var(--ink);}

/* 자동 저장용 보이지 않는 조각 */
.st-key-store_box{height:0; margin:0 0 -1rem; padding:0; overflow:hidden;}

/* 메모·별표 */
.wcard{position:relative;}
.wcard .star{position:absolute; top:4px; right:8px; color:#D99A00; font-size:1.05rem;}
.wcard .note{font-size:.78rem; color:var(--soft); margin-top:.3rem; text-align:left; padding:0 .4rem; overflow:hidden;
  display:-webkit-box; -webkit-line-clamp:3; -webkit-box-orient:vertical; word-break:break-all;}
.memo{max-width:520px; margin:.4rem auto .8rem; padding:.6rem .9rem; background:#FFF7E0; color:var(--ink);
  border-left:4px solid var(--seal); border-radius:6px; white-space:pre-wrap; font-size:.98rem;}

/* 사자성어 고르기 버튼 (흰 바탕 + 어두운 글씨) */
div[class*="st-key-xc_"] button{min-height:4rem; background:#fff; border:1.5px solid var(--line);}
div[class*="st-key-xc_"] button, div[class*="st-key-xc_"] button p{color:var(--ink);}
div[class*="st-key-xc_text"] button p{font-size:1.1rem;}
div[class*="st-key-xc_hj"] button p, div[class*="st-key-xc_blank"] button p{font-family:var(--serif); font-size:2rem; line-height:1.15;}

/* 탭이 많아서 모바일에서는 조금 좁게 */
@media (max-width:640px){
  .stTabs [data-baseweb="tab"]{padding:0 .6rem; font-size:.92rem; height:2.9rem;}
}

/* 사자성어 목록 */
.idm{background:#fff; border:1px solid var(--line); border-radius:10px; padding:.55rem .8rem; margin:.4rem 0; color:var(--ink);}
.idm .idm-h{font:700 1.45rem/1.3 var(--serif); letter-spacing:.06em;}
.idm .idm-s{color:var(--indigo); font-weight:600; margin-left:.4rem;}
.idm .idm-m{color:var(--soft); font-size:.92rem; margin-top:.15rem; line-height:1.4;}
.idm summary{cursor:pointer; list-style:none;}
.idm .miss{color:#C0392B; border-bottom:2px solid #C0392B;}
details.idm[open] summary{margin-bottom:.3rem;}

/* 시작 화면 3단계 */
.steps{display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin:.4rem 0 1rem;}
.step{display:flex; flex-direction:column; align-items:center; gap:.4rem; text-align:center; padding:.8rem .4rem;
  background:#fff; border:1px solid var(--line); border-radius:10px; color:var(--ink); font-size:.92rem; line-height:1.35;}
.step b{display:grid; place-items:center; width:30px; height:30px; border-radius:50%; background:var(--indigo);
  color:#fff; font:700 1rem/1 var(--serif);}

/* 한자 분해(구성 글자) */
.wcard details.bu summary{cursor:pointer; list-style:none; padding:0; font:500 2.6rem/1.1 var(--serif); color:var(--ink);}
.wcard details.bu summary::-webkit-details-marker{display:none;}
.wcard details.bu summary::after{content:"🧩"; font-size:.65rem; margin-left:.1rem; opacity:.55; vertical-align:super;}
.wcard:has(details.bu[open]){grid-column:1/-1;}
.bu-body{display:none; margin-top:.6rem; padding:.6rem .75rem; background:#F6F8FC; border:1px solid var(--line);
  border-radius:8px; text-align:left; color:var(--ink); font-size:.95rem;}
.wcard details.bu[open] ~ .bu-body, .bu-body.open{display:block;}
.bu-char{padding:.45rem 0; border-top:1px dashed var(--line);}
.bu-char:first-child{border-top:0;}
.bu-eq{display:flex; flex-wrap:wrap; align-items:center; gap:.35rem .5rem;}
.bu-big{font:700 2.2rem/1 var(--serif); color:var(--seal);}
.bu-op{color:var(--soft); font-weight:700;}
.bu-part{display:inline-flex; flex-direction:column; align-items:center; min-width:3.4rem; padding:.25rem .5rem;
  background:#fff; border:1px solid var(--line); border-radius:8px; line-height:1.2;}
.bu-part b{font:500 1.7rem/1.1 var(--serif);}
.bu-part i{font-style:normal; font-size:.78rem; color:var(--soft);}
.bu-part em{font-style:normal; font-size:.72rem; color:var(--seal); font-weight:700;}
.bu-part.rad{border-color:var(--jade); background:#EEF7F4;}
.bu-note{color:var(--soft); font-size:.82rem; margin-top:.2rem;}
.bu-sub{margin:.4rem 0 0 .3rem; padding-left:.7rem; border-left:3px solid var(--line);}
</style>
"""


def inject_css():
    st.markdown(CSS, unsafe_allow_html=True)


def html(s):
    st.markdown(s, unsafe_allow_html=True)


# ----------------------------------------------------------------------
# 공통 조각
# ----------------------------------------------------------------------
def hanja_card(text):
    n = len(text)
    size = {1: 128, 2: 92, 3: 70}.get(n, 54)
    cls = "hj-card" if n == 1 else "hj-card wide"
    html(f"<div class='{cls}' style='font-size:min({size}px,24vw)'>{esc(text)}</div>")


def meaning_card(w):
    html(f"<div class='mean-card'>{esc(w['meaning'])}<b>{esc(w['sound'])}</b></div>")


def reveal(w):
    html(
        f"<div class='reveal'><div class='rh'>{esc(w['hanja'])}</div>"
        f"<div class='rm'>{esc(w['meaning'])} <b>{esc(w['sound'])}</b></div></div>"
    )


def verdict_row(ok, w, extra=""):
    cls, txt = ("ok", "정답") if ok else ("no", "오답")
    html(
        f"<div class='fb-row'><div class='verdict {cls}'>{txt}</div>"
        f"<div class='reveal'><div class='rh'>{esc(w['hanja'])}</div>"
        f"<div class='rm'>{esc(w['meaning'])} <b>{esc(w['sound'])}</b>{extra}</div></div></div>"
    )


def streak_now(results):
    n = 0
    for r in reversed(results):
        if not r["correct"]:
            break
        n += 1
    return n


# ----------------------------------------------------------------------
# 화면: 머리글 / 시작 화면
# ----------------------------------------------------------------------
def header():
    s = st.session_state
    playing = bool(s.quiz) and s.pos < len(s.quiz)  # 문제 푸는 중엔 머리글을 줄여서 화면을 아껴요
    sub = "" if playing else "<div class='sub'>쓰고, 고르고, 넘겨 보며 익혀요</div>"
    html(
        f"<div class='app-head{' compact' if playing else ''}'><div class='seal'>漢</div>"
        f"<div><div class='title'>한자 연습장</div>{sub}</div></div>"
    )
    if not s.words or playing:
        return
    new, review = due_counts(s.words)
    goal = max(1, int(s.get("set_goal", 20)))
    done = hist_today()
    chips = [f"<span class='chip'>📚 한자 {len(s.words)}개</span>"]
    days = study_streak()
    if days >= 2:
        chips.append(f"<span class='chip good'>🔥 {days}일 연속 학습</span>")
    if review:
        chips.append(f"<span class='chip hot'>🔁 복습할 차례 {review}개</span>")
    if new:
        chips.append(f"<span class='chip'>🌱 아직 안 푼 {new}개</span>")
    if s.wrong:
        chips.append(f"<span class='chip'>📕 오답 노트 {len(s.wrong)}개</span>")
    dd = exam_dday()
    if dd is not None and dd >= 0:
        chips.append(f"<span class='chip hot'>📅 시험 {'D-Day' if dd == 0 else f'D-{dd}'}</span>")
    if s.store_ok:
        chips.append("<span class='chip good'>💾 자동 저장 켜짐</span>")
    elif s.store_ok is False and s.restored:
        chips.append("<span class='chip hot'>⚠️ 이 브라우저에선 자동 저장이 안 돼요</span>")
    html("<div class='chips'>" + "".join(chips) + "</div>")
    st.progress(min(1.0, done / goal), text=f"오늘의 목표 {done} / {goal}문제" + (" 🎉 달성!" if done >= goal else ""))


SAMPLE_CSV = (
    "한자,뜻,음,분류\n"
    "天,하늘,천,자연\n地,땅,지,자연\n山,메,산,자연\n水,물,수,자연\n火,불,화,자연\n"
    "日,날,일,자연\n月,달,월,자연\n木,나무,목,자연\n金,쇠,금,자연\n土,흙,토,자연\n"
    "一,한,일,숫자\n二,두,이,숫자\n三,석,삼,숫자\n四,넉,사,숫자\n五,다섯,오,숫자\n六,여섯,육/륙,숫자\n"
    "人,사람,인,사람\n女,계집,녀/여,사람\n男,사내,남,사람\n子,아들,자,사람\n父,아비,부,사람\n母,어미,모,사람\n"
)


def load_panel():
    """단어장 불러오기 (파일 업로드 / 직접 붙여넣기). 단어장이 없을 땐 시작 화면에, 있을 땐 사이드바에 나와요"""
    s = st.session_state
    f = st.file_uploader("한자 파일 (txt / csv / xlsx)", key=f"wordfile_{s.up_gen}")  # 형식 제한 없음: 태블릿 호환
    st.download_button("샘플 CSV 받기", SAMPLE_CSV.encode("utf-8-sig"), "sample_hanja.csv", "text/csv")
    st.caption("한 줄에 '한자, 뜻, 음'. 4번째 열에 '급수/단원/분류'를 넣으면 범위를 골라 풀 수 있어요.")
    if f is not None:
        sig = (f.name, f.size)
        if sig != s.file_sig:
            try:
                words, dropped, merged = parse_file(f)
                s.file_sig = sig
                if words:
                    s.words = words
                    s.quiz, s.cards = [], None
                    s.flash = _loaded_msg(len(words), dropped, merged)
                else:
                    s.flash = "읽을 수 있는 한자가 없어요. 한 줄에 '한자 뜻 음' 모양인지 확인해 주세요."
            except Exception as e:
                s.flash = f"파일을 읽지 못했어요: {e}"
    with st.expander("✍️ 텍스트 직접 붙여넣기 (파일이 안 될 때)"):
        pasted = st.text_area("한 줄에 '한자 뜻 음' (예: 天 하늘 천)", height=150, key="paste_text")
        if s.words:
            st.button("➕ 지금 단어장에 추가", type="primary", on_click=paste_add_cb, args=("paste_text",),
                      key="paste_add_btn", width="stretch",
                      help="지금 단어장은 그대로 두고, 붙여넣은 한자를 뒤에 더해요. 이미 있는 한자는 새 뜻·음만 덧붙여요.")
            st.caption("아래 버튼은 지금 단어장을 지우고 붙여넣은 내용으로 바꿔요.")
        if st.button("붙여넣은 내용으로 교체" if s.words else "붙여넣은 내용 적용", key="paste_apply_btn"):
            words, dropped, merged = build_words(parse_text(strip_comments(pasted)))
            if words:
                s.words, s.quiz, s.cards = words, [], None
                s.file_sig = (f.name, f.size) if f is not None else None  # 업로드 파일이 덮어쓰지 않게
                s.flash = _loaded_msg(len(words), dropped, merged)
            else:
                s.flash = "읽을 수 있는 줄이 없어요. 한 줄에 '한자 뜻 음' 모양인지 확인해 주세요."


def load_sample():
    s = st.session_state
    df = pd.read_csv(io.StringIO(SAMPLE_CSV), header=None, dtype=str, keep_default_na=False)
    words, _, _ = build_words(parse_table(df))
    s.words, s.quiz, s.cards = words, [], None
    f = s.get(f"wordfile_{s.up_gen}")
    s.file_sig = (f.name, f.size) if f is not None else None  # 이미 올려 둔 파일이 샘플을 덮어쓰지 않게
    s.flash = f"샘플 한자 {len(words)}개를 불러왔어요. 마음껏 둘러보세요!"


def render_welcome():
    st.subheader("단어장을 불러와 볼까요?")
    html(
        "<div class='steps'>"
        "<div class='step'><b>1</b><span>한자·뜻·음이 적힌<br>파일을 올려요</span></div>"
        "<div class='step'><b>2</b><span>문제 · 카드로<br>연습해요</span></div>"
        "<div class='step'><b>3</b><span>매일 조금씩<br>복습해요</span></div>"
        "</div>"
    )
    st.button("✨ 샘플로 먼저 둘러보기", on_click=load_sample, width="stretch",
              help="파일이 없어도 샘플 한자 22개로 바로 체험해 볼 수 있어요.")
    set_picker("word", "welcome")
    load_panel()
    st.caption("🈴 단어장을 불러오면, 그 한자들로만 이루어진 사자성어가 '사자성어' 탭에 자동으로 나와요.")
    st.caption("💾 푼 기록과 단어장은 이 기기의 브라우저에 자동으로 저장돼서, 다시 들어오면 이어서 할 수 있어요. "
               "(다른 기기와는 따로 저장돼요)")


# ----------------------------------------------------------------------
# 화면: 연습 (설정 → 문제 → 결과)
# ----------------------------------------------------------------------
def quick_start():
    """복습할 차례인 한자부터 바로 시작"""
    st.session_state.set_source = "오늘 복습할 것"
    start_quiz()


def render_home():
    s = st.session_state
    render_plan_card()
    st.subheader("오늘은 무엇을 연습할까요?")
    new, review = due_counts(s.words)
    if review + new:
        st.button(f"⚡ 오늘 복습 바로 시작 (복습 {review} · 새 한자 {new})", on_click=quick_start, width="stretch")
    st.radio("문제 유형", ["섞어서", "뜻·음 쓰기", "한자 쓰기", "객관식"], key="set_type", horizontal=True,
             help="뜻·음 쓰기: 키보드로 입력 / 한자 쓰기: 손글씨 / 객관식: 4개 중 고르기")
    st.radio("문제 범위", SOURCES, key="set_source", horizontal=True,
             help=f"'오늘 복습할 것'은 복습 날짜가 된 한자({review}개)와 아직 안 푼 한자({new}개)를 약한 것부터 골라줘요.")
    groups = sorted({w["group"] for w in s.words if w.get("group")},
                    key=lambda g: [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", g)])
    if groups:
        s.set_groups = [g for g in s.set_groups if g in groups]
        st.multiselect("급수/단원/분류 (비워 두면 전체)", groups, key="set_groups", placeholder="전체 (눌러서 고르기)")
    c1, c2 = st.columns([3, 2])
    c1.slider("문제 수", 5, 100, step=5, key="set_count")
    with c2:
        st.write("")
        st.toggle("순서 섞기", key="set_shuffle")
    st.button("▶️ 시작하기", type="primary", on_click=start_quiz, width="stretch")
    if s.get("set_type") in ("한자 쓰기", "섞어서") and not find_font():
        st.caption("한자 폰트를 찾지 못해 손글씨는 자가 채점으로 진행돼요.")


def render_practice():
    s = st.session_state
    if not s.quiz:
        render_home()
        return
    if s.pos >= len(s.quiz):
        render_summary()
        return
    q = s.quiz[s.pos]
    st.progress(s.pos / len(s.quiz), text=f"{s.pos + 1} / {len(s.quiz)}  ·  {TYPE_LABELS[q['type']]}")
    streak = streak_now(s.results)
    chips = [f"<span class='chip'>현재 점수 {sum(r['correct'] for r in s.results)} / {len(s.results)}</span>"]
    if streak >= 2:
        chips.append(f"<span class='chip good'>🔥 {streak}연속 정답</span>")
    html("<div class='chips'>" + "".join(chips) + "</div>")
    if s.feedback:
        render_feedback(q)
    elif q["type"] == "read":
        render_read(q["w"])
    elif q["type"] == "write":
        render_write(q["w"])
    else:
        render_choice(q)
    if not s.feedback:
        st.button("그만하기", key="quit_quiz", on_click=reset_quiz)


# --- 뜻·음 쓰기 ---------------------------------------------------------------
def render_read(w):
    s = st.session_state
    hanja_card(w["hanja"])
    html("<div class='ask'>이 한자의 <b>뜻과 음</b>을 입력하세요. 예) 하늘 천</div>")
    if s.hint:
        first_mean = split_alts(w["meaning"])[0] if split_alts(w["meaning"]) else ""
        st.caption(f"💡 음은 '{w['sound'][0]}'(으)로 시작하고, 뜻은 {len(first_mean)}글자예요.")
    c1, c2, _ = st.columns([1, 1, 2])
    if c1.button("💡 힌트", width="stretch"):
        s.hint = True
        st.rerun()
    if c2.button("모르겠어요", key="skip_read", width="stretch"):
        finish_read(w, "")
    ans = st.chat_input("뜻과 음을 입력하세요 (예: 하늘 천)")
    if ans:
        finish_read(w, ans)


def finish_read(w, ans):
    s = st.session_state
    ok, detail = check_reading(ans, w) if ans else (False, "")
    record(w, "read", ok, ans)
    s.feedback = {"type": "read", "user": ans or "(모르겠어요)", "ok": ok, "detail": detail}
    st.rerun()


# --- 객관식 -------------------------------------------------------------------
def render_choice(q):
    s = st.session_state
    w, ty = q["w"], q["type"]
    if ty == "choice_ko":
        hanja_card(w["hanja"])
        html("<div class='ask'>이 한자의 <b>뜻과 음</b>은?</div>")
    else:
        meaning_card(w)
        html("<div class='ask'>이 뜻·음에 해당하는 <b>한자</b>는?</div>")
    with st.container(key="choices_ko" if ty == "choice_ko" else "choices_hj"):
        cols = st.columns(2)
        for i, o in enumerate(q["opts"]):
            text = o["label"] if ty == "choice_ko" else o["hanja"]
            cols[i % 2].button(text, key=f"opt_{s.pos}_{s.nonce}_{i}", on_click=answer_choice,
                               args=(w, ty, o), width="stretch")
    st.button("모르겠어요", key="skip_choice", on_click=answer_choice, args=(w, ty, None))


def answer_choice(w, ty, o):
    s = st.session_state
    ok = o is not None and o["hanja"] == w["hanja"]
    if o is None:
        shown, detail = "(모르겠어요)", ""
    else:
        shown = o["hanja"] if ty == "choice_hj" else o["label"]
        detail = ""
        if not ok:
            detail = (f"고른 '{o['label']}'은(는) {o['hanja']}의 뜻·음이에요." if ty == "choice_ko"
                      else f"고른 {o['hanja']}은(는) '{o['label']}'이에요.")
    record(w, ty, ok, shown)
    s.feedback = {"type": "read", "user": shown, "ok": ok, "detail": detail, "choice": True}


# --- 한자 쓰기 ----------------------------------------------------------------
def canvas_px():
    """한자를 쓰는 칸의 한 변 길이(px). 사이드바의 '쓰는 칸 크기'에서 바꿀 수 있어요"""
    return _int(st.session_state.get("set_canvas_px"), 180, 120, 360)


def flatten(data):
    """캔버스 이미지(RGBA)를 흰 바탕에 합성해요. 배경이 투명해도 빈 곳이 '검은 획'으로 읽히지 않게 하는 안전장치"""
    a = data[:, :, 3:4].astype(float) / 255.0
    rgb = data[:, :, :3].astype(float) * a + 255.0 * (1.0 - a)
    out = np.dstack([rgb, np.full(data.shape[:2], 255.0)])
    return out.astype(np.uint8)


def canvas_backdrop(img, px):
    """캔버스 뒤쪽 칸(iframe)에 십자 점선·힌트 글씨를 CSS 배경으로 깔아요. 캔버스 자체는 투명해서 채점에는 영향이 없어요"""
    buf = io.BytesIO()
    img.save(buf, "PNG")
    uri = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
    html(
        f"<style>.st-key-canvas_box{{width:{px + 52}px; max-width:100%; background:url(" + uri + ") 0 0 / "
        f"{px}px {px}px no-repeat;}}"
        ".st-key-canvas_box iframe{mix-blend-mode:multiply;}</style>"
    )


def draw_canvas(key, px):
    s = st.session_state
    kwargs = dict(
        fill_color="rgba(255,255,255,0)",
        stroke_width=int(s.get("set_stroke", 8)),
        stroke_color="#000000",
        background_color="",
        height=px,
        width=px,
        drawing_mode="freedraw",
        key=key,
    )
    try:  # 새 버전 캔버스 라이브러리는 이 옵션을 켜야 글씨 이미지를 돌려줘요 (옛 버전엔 없는 옵션)
        if "return_image_data" in inspect.signature(st_canvas).parameters:
            kwargs["return_image_data"] = True
    except Exception:
        pass
    try:
        return st_canvas(**kwargs)
    except Exception as e:
        st.error(f"캔버스를 표시하지 못했어요: {type(e).__name__}: {e}")
        return None


def render_write(w):
    s = st.session_state
    font = find_font()
    px = canvas_px()
    left, right = st.columns([3, 2])  # 쓰는 칸이 작아졌으니 오른쪽 칸도 좁게
    with left:
        meaning_card(w)
        html("<div class='ask'>에 해당하는 한자를 오른쪽 칸에 쓰세요</div>")
        hint = st.checkbox("연한 글씨 가이드 보기 (힌트)", key=f"hint_{s.pos}_{s.nonce}", disabled=font is None)
        if font is None:  # 체크박스가 회색으로 막혀 있는 이유를 알려줘요
            st.caption("⚠️ 서버에서 한자 폰트를 불러오지 못해 힌트를 쓸 수 없어요. 제출하면 정답과 비교해서 직접 채점해요.")
    # 십자 점선은 항상 깔고, 힌트를 켜면 정답 글자가 연한 회색으로 같이 나와요
    canvas_backdrop(guide_image(w["hanja"], font, px) if (hint and font) else grid_image(px), px)
    with right:
        with st.container(key="canvas_box"):
            res = draw_canvas(f"cv_{s.pos}_{s.nonce}_{int(hint)}_{px}", px)
    b1, b2, b3 = st.columns(3)
    if b1.button("제출", type="primary", width="stretch"):
        data = flatten(res.image_data) if (res is not None and res.image_data is not None) else None
        if data is None or (data[:, :, :3].astype(float).mean(axis=2) < 110).sum() < 30:
            st.warning("먼저 한자를 써 주세요.")
        else:
            submit_write(w, data, font)
    if b2.button("지우기", width="stretch"):
        s.nonce += 1
        st.rerun()
    if b3.button("모르겠어요", key="skip_write", width="stretch"):
        record(w, "write", False, "")
        s.feedback = {"type": "write", "ok": False, "pending": False, "score": None, "overlay": None}
        st.rerun()


def submit_write(w, data, font):
    s = st.session_state
    if font:
        cands = [x["hanja"] for x in s.words if len(x["hanja"]) == 1]
        with st.spinner("채점 중..."):
            r = grade_writing(data, w["hanja"], font, grade_params(), cands)
        if r is None:
            st.warning("글씨를 인식하지 못했어요. 조금 더 크게 써 주세요.")
            return
        record(w, "write", r["ok"], f"유사도 {r['score']:.2f}")
        s.feedback = {
            "type": "write", "ok": r["ok"], "pending": False, "score": r["score"], "overlay": r["overlay"],
            "detail": r["detail"], "reason": r["reason"], "p": r["p"], "r": r["r"],
            "gap": max(r["miss"], r["extra"]), "other": r["other"],
        }
    else:  # 폰트가 없으면 자가 채점
        s.feedback = {"type": "write", "ok": False, "pending": True, "score": None, "overlay": None}
    st.rerun()


def self_grade(w, ok):
    s = st.session_state
    record(w, "write", ok, "자가 채점")
    s.feedback["pending"] = False
    s.feedback["ok"] = ok
    st.rerun()


# --- 결과 화면 ----------------------------------------------------------------
def render_feedback(q):
    s = st.session_state
    fb, w = s.feedback, q["w"]
    if fb["type"] == "read":
        with st.chat_message("user"):
            st.write(fb["user"])
        verdict_row(fb["ok"], w)
        if fb.get("detail"):
            st.caption(fb["detail"])
    else:
        if fb["pending"]:
            hanja_card(w["hanja"])
            html(f"<div class='ask'>{esc(w['meaning'])} {esc(w['sound'])}</div>")
            st.info("정답 한자와 내 글씨를 비교해서 직접 채점해 주세요. (서버에 한자 폰트가 없어 자동 채점이 꺼져 있어요)")
            g1, g2, _ = st.columns([1, 1, 2])
            if g1.button("⭕ 맞았어요", width="stretch"):
                self_grade(w, True)
            if g2.button("❌ 틀렸어요", width="stretch"):
                self_grade(w, False)
            return
        extra = f"<span style='color:#5B6479;font-size:.9rem'> · 유사도 {fb['score']:.2f}</span>" if fb["score"] is not None else ""
        verdict_row(fb["ok"], w, extra)
        if fb.get("overlay") is not None:
            st.image(fb["overlay"], caption="빨강: 정답 글자 / 파랑: 내 글씨 / 보라: 겹친 부분", width=256)
        if fb.get("detail"):
            st.caption(fb["detail"])
        if fb.get("p") is not None:
            with st.expander("채점 자세히 보기"):
                line = (f"내 글씨가 정답과 맞는 정도 {fb['p']:.0%} · 정답 획을 채운 정도 {fb['r']:.0%}"
                        f" · 가장 크게 어긋난 획 {fb['gap']:.0f}%")
                if fb.get("other"):
                    line += f" · 가장 비슷한 다른 글자 '{fb['other'][0]}' {fb['other'][1]:.2f}"
                st.caption(line)
    if not fb["ok"] and not fb.get("pending"):
        show_memo(w["hanja"])
    if not fb["ok"] and not fb.get("pending") and not fb.get("choice"):
        st.button("🙋 내가 맞게 썼어요 (정정)", on_click=override_correct)
    st.button("다음 문제 ➡️", type="primary", on_click=next_question, width="stretch")


def render_summary():
    s = st.session_state
    res = s.results
    n = len(res)
    c = sum(r["correct"] for r in res)
    best, cur = 0, 0
    for r in res:
        cur = cur + 1 if r["correct"] else 0
        best = max(best, cur)
    st.subheader("수고했어요! 결과예요")
    rate = c / n * 100 if n else 0
    html(f"<div class='score-box'><div class='big'>{c} / {n}</div><div class='lbl'>정답률 {rate:.0f}%</div></div>")
    m1, m2 = st.columns(2)
    m1.metric("최고 연속 정답", f"{best}개")
    m2.metric("오답 노트 누적", len(s.wrong))
    if n:
        df = pd.DataFrame(
            [
                {
                    "결과": "⭕" if r["correct"] else "❌",
                    "유형": TYPE_LABELS.get(r["type"], r["type"]),
                    "한자": r["w"]["hanja"],
                    "뜻": r["w"]["meaning"],
                    "음": r["w"]["sound"],
                    "내 답": r["answer"],
                }
                for r in res
            ]
        )
        st.dataframe(df, width="stretch", hide_index=True)
    wrong_items = [make_item(r["w"], r["type"]) for r in res if not r["correct"]]
    b1, b2 = st.columns(2)
    b1.button("🔁 새로 시작", on_click=start_quiz, width="stretch")
    if wrong_items:
        b2.button("❌ 이번에 틀린 것만 다시", on_click=reset_round, args=(wrong_items, False), width="stretch")
    st.button("설정 바꾸기", on_click=reset_quiz)


# ----------------------------------------------------------------------
# 화면: 카드
# ----------------------------------------------------------------------
def render_cards():
    s = st.session_state
    c = s.cards
    if c is None:
        st.subheader("카드로 넘겨 보며 익혀요")
        st.write("카드를 보고 떠올린 뒤 뒤집어서 확인하세요. '알아요/헷갈려요'는 복습 주기에 반영돼요.")
        st.radio("카드 앞면", ["한자 → 뜻·음", "뜻·음 → 한자"], key="card_dir", horizontal=True)
        st.radio("카드 범위", SOURCES, key="card_src", horizontal=True)
        a, b = st.columns([3, 2])
        a.slider("카드 수", 5, 100, step=5, key="card_count")
        with b:
            st.write("")
            st.toggle("순서 섞기", key="card_shuffle")
        st.button("🃏 카드 시작", type="primary", on_click=start_cards, width="stretch")
        return
    deck, i = c["deck"], c["i"]
    if i >= len(deck):
        st.subheader("카드를 다 봤어요")
        k, u = len(c["known"]), len(c["unsure"])
        html(f"<div class='chips'><span class='chip good'>알아요 {k}개</span><span class='chip hot'>헷갈려요 {u}개</span></div>")
        if u:
            html("<div class='ask'>헷갈린 한자</div>")
            html("<div class='wgrid'>" + "".join(
                f"<div class='wcard'><div class='h'>{esc(w['hanja'])}</div><div class='m'><b>{esc(w['meaning'])}</b> {esc(w['sound'])}</div></div>"
                for w in c["unsure"]) + "</div>")
        b1, b2 = st.columns(2)
        if u:
            b1.button("🔁 헷갈린 것만 다시", on_click=card_retry_unsure, width="stretch")
        b2.button("끝내기", on_click=card_close, width="stretch")
        return
    cur = deck[i]
    st.progress(i / len(deck), text=f"{i + 1} / {len(deck)}")
    front_is_hanja = s.get("card_dir", "한자 → 뜻·음") == "한자 → 뜻·음"
    if not c["flip"]:
        hanja_card(cur["hanja"]) if front_is_hanja else meaning_card(cur)
        html("<div class='ask'>떠올려 보고 뒤집어 확인하세요</div>")
        st.button("🔄 뒤집기", type="primary", on_click=card_flip, width="stretch", key="card_flip")
    else:
        hanja_card(cur["hanja"])
        verdict_row_neutral(cur)
        show_memo(cur["hanja"])
        b1, b2 = st.columns(2)
        b1.button("😊 알아요", on_click=card_mark, args=(True,), width="stretch", type="primary", key="card_known")
        b2.button("🤔 헷갈려요", on_click=card_mark, args=(False,), width="stretch", key="card_unsure")
    st.button("그만 보기", key="card_quit", on_click=card_close)


def verdict_row_neutral(w):
    html(f"<div class='fb-row'>{reveal_html(w)}</div>")


def reveal_html(w):
    return (f"<div class='reveal'><div class='rh'>{esc(w['hanja'])}</div>"
            f"<div class='rm'>{esc(w['meaning'])} <b>{esc(w['sound'])}</b></div></div>")


# ----------------------------------------------------------------------
# 화면: 오답 노트 / 기록 / 단어장
# ----------------------------------------------------------------------
def render_wrong():
    s = st.session_state
    if not s.wrong:
        st.info("오답 노트가 비어 있어요. 틀린 문제가 여기에 쌓여요.")
        return
    df = pd.DataFrame(
        [
            {"한자": v["hanja"], "뜻": v["meaning"], "음": v["sound"], "틀린 횟수": v["wrong_count"], "최근": v["last"]}
            for v in s.wrong.values()
        ]
    ).sort_values("틀린 횟수", ascending=False)
    st.dataframe(df, width="stretch", hide_index=True)
    sel = st.multiselect("오답 노트에서 뺄 한자", list(s.wrong.keys()))
    c1, c2, c3 = st.columns(3)
    if c1.button("선택 삭제", width="stretch") and sel:
        for k in sel:
            s.wrong.pop(k, None)
        st.rerun()
    if c2.button("전체 비우기", width="stretch"):
        s.wrong.clear()
        st.rerun()
    c3.download_button("CSV로 저장", df.to_csv(index=False).encode("utf-8-sig"), "오답노트.csv", "text/csv", width="stretch")

    def go_wrong():
        s.set_source = "오답 노트만"
        s.set_type = "섞어서"
        start_quiz()

    st.button("📝 오답으로 바로 풀기", type="primary", on_click=go_wrong, width="stretch")
    st.caption("오답 노트 모드에서 맞히면 노트에서 자동으로 빠져요. 풀이는 '연습' 탭에서 진행돼요.")


# ----------------------------------------------------------------------
# 학습 잔디 (깃허브 스타일 히트맵) — 최근 1년, 하루에 푼 문제가 많을수록 진해져요
# ----------------------------------------------------------------------
HM_COLORS = ["#E7EBF2", "#BFE0D6", "#86C5B4", "#4FA08D", "#2F7D6D"]
HM_WEEKS = 53


def _hm_day_totals(history):
    """{날짜: (맞음, 틀림)} — 날짜가 이상한 기록은 건너뛰어요"""
    out = {}
    for k, v in (history or {}).items():
        try:
            d = date.fromisoformat(str(k))
            r, w = int(v.get("right", 0)), int(v.get("wrong", 0))
        except Exception:
            continue
        if r + w > 0:
            out[d] = (r, w)
    return out


def study_streaks(history, today):
    """(지금 연속일, 최장 연속일, 최근 1년 학습한 날 수). 오늘 아직 안 풀었어도 어제까지 이어졌으면 연속이 유지돼요"""
    days = _hm_day_totals(history)
    if not days:
        return 0, 0, 0
    cur, d = 0, today if today in days else today - timedelta(days=1)
    while d in days:
        cur += 1
        d -= timedelta(days=1)
    best = run = 0
    prev = None
    for d in sorted(days):
        run = run + 1 if prev is not None and (d - prev).days == 1 else 1
        best = max(best, run)
        prev = d
    year_ago = today - timedelta(days=364)
    return cur, best, sum(1 for d in days if year_ago <= d <= today)


def heatmap_levels(totals):
    """하루 학습량 → 0~4단계. 가장 많이 푼 날 하나에 쏠리지 않게 사분위수로 나눠요"""
    vals = sorted(t for t in totals if t > 0)
    if not vals:
        return lambda t: 0
    q1, q2, q3 = vals[(len(vals) - 1) // 4], vals[(len(vals) - 1) // 2], vals[(3 * (len(vals) - 1)) // 4]
    return lambda t: 0 if t <= 0 else 1 + (t > q1) + (t > q2) + (t > q3)


def heatmap_html(history, today):
    """iframe 안에 넣을 완성된 HTML (오른쪽 끝 = 이번 주로 자동 스크롤)"""
    days = _hm_day_totals(history)
    start = today - timedelta(days=(today.weekday() + 1) % 7) - timedelta(weeks=HM_WEEKS - 1)  # 일요일 시작
    span = [start + timedelta(days=i) for i in range(HM_WEEKS * 7)]
    level = heatmap_levels([sum(days.get(d, (0, 0))) for d in span if d <= today])
    cells = []
    for w in range(HM_WEEKS):
        week = span[w * 7:(w + 1) * 7]
        first = next((d for d in week if d.day == 1 and d <= today), None)
        cells.append(f"<div class='mo'>{first.month}월</div>" if first else "<div class='mo'></div>")
        for d in week:
            if d > today:
                cells.append("<div class='c x'></div>")
                continue
            r, wr = days.get(d, (0, 0))
            tip = f"{d.year}.{d.month}.{d.day} · " + (f"맞음 {r} · 틀림 {wr}" if r + wr else "푼 문제 없음")
            cls = f"c l{level(r + wr)}" + (" t" if d == today else "")
            cells.append(f"<div class='{cls}' title='{esc(tip)}' data-tip='{esc(tip)}'></div>")
    daylab = "".join(f"<div>{x}</div>" for x in ["", "", "월", "", "수", "", "금", ""])
    colors = "".join(f".l{i}{{background:{c}}}" for i, c in enumerate(HM_COLORS))
    legend = "".join(f"<span class='c l{i}'></span>" for i in range(5))
    return f"""<!doctype html><html><head><meta charset='utf-8'><style>
body{{margin:0; font-family:'Noto Sans KR','Apple SD Gothic Neo','Malgun Gothic',system-ui,sans-serif; color:#1E2438;}}
#w{{overflow-x:auto; padding:2px 2px 6px;}}
.in{{display:flex; width:max-content;}}
.dl{{display:grid; grid-template-rows:repeat(8,12px); gap:3px; font-size:10px; color:#5B6479; margin-right:6px; line-height:12px; position:sticky; left:0; background:#fff; z-index:1; padding-right:2px;}}
.g{{display:grid; grid-template-rows:repeat(8,12px); grid-auto-flow:column; grid-auto-columns:12px; gap:3px;}}
.mo{{font-size:10px; color:#5B6479; line-height:12px; white-space:nowrap; overflow:visible;}}
.c{{width:12px; height:12px; border-radius:3px; box-sizing:border-box;}}
.c.x{{background:transparent;}}
.c.t{{outline:2px solid #B8372C; outline-offset:-1px;}}
{colors}
#i{{font-size:12px; color:#5B6479; margin-top:4px; min-height:16px;}}
.lg{{display:flex; align-items:center; gap:3px; font-size:11px; color:#5B6479; justify-content:flex-end; margin-top:4px;}}
</style></head><body>
<div id='w'><div class='in'><div class='dl'>{daylab}</div><div class='g'>{''.join(cells)}</div></div></div>
<div class='lg'>적음 {legend} 많음</div>
<div id='i'>칸을 누르면 그날 기록이 나와요.</div>
<script>
var w=document.getElementById('w'); w.scrollLeft=w.scrollWidth;
document.querySelectorAll('.c[data-tip]').forEach(function(e){{e.addEventListener('click',function(){{document.getElementById('i').textContent=e.getAttribute('data-tip');}});}});
</script></body></html>"""


def render_heatmap():
    s = st.session_state
    today = today_date()
    cur, best, n_year = study_streaks(s.history, today)
    st.markdown("##### 🌱 학습 잔디 (최근 1년)")
    a, b, c = st.columns(3)
    a.metric("🔥 지금 연속", f"{cur}일")
    b.metric("🏆 최장 연속", f"{best}일")
    c.metric("📅 1년간 공부한 날", f"{n_year}일")
    components.html(heatmap_html(s.history, today), height=170)
    st.caption("하루에 푼 문제가 많을수록 진한 초록이에요. 빨간 테두리가 오늘이에요. 오늘 아직 안 풀었어도 어제까지 이어졌다면 연속 기록은 유지돼요.")


def render_stats():
    s = st.session_state
    if not s.stats:
        st.info("아직 푼 문제가 없어요.")
        return
    rows = []
    for k, v in s.stats.items():
        tot = v["right"] + v["wrong"]
        rows.append(
            {"한자": k, "뜻": v["meaning"], "음": v["sound"], "맞음": v["right"], "틀림": v["wrong"],
             "정답률(%)": round(v["right"] / tot * 100) if tot else 0}
        )
    df = pd.DataFrame(rows).sort_values("정답률(%)")
    tr, tw = int(df["맞음"].sum()), int(df["틀림"].sum())
    a, b, c = st.columns(3)
    a.metric("푼 문제", tr + tw)
    b.metric("전체 정답률", f"{tr / (tr + tw) * 100:.0f}%" if tr + tw else "-")
    c.metric("약한 한자 (정답률 50% 미만)", int((df["정답률(%)"] < 50).sum()))

    if s.history:
        render_heatmap()
        with st.expander("최근 14일 맞음·틀림 막대그래프"):
            days = [(today_date() - timedelta(days=d)).isoformat() for d in range(13, -1, -1)]
            hd = pd.DataFrame(
                {"맞음": [s.history.get(d, {}).get("right", 0) for d in days],
                 "틀림": [s.history.get(d, {}).get("wrong", 0) for d in days]},
                index=[d[5:] for d in days],
            )
            st.bar_chart(hd, color=["#2F7D6D", "#B8372C"], sort=False)

    weak = df[df["틀림"] > 0].sort_values(["정답률(%)", "틀림"], ascending=[True, False]).head(10)
    if len(weak):
        st.markdown("##### 가장 약한 한자 TOP 10")
        html("<div class='wgrid'>" + "".join(
            f"<div class='wcard'><div class='h'>{esc(r['한자'])}</div>"
            f"<div class='m'><b>{esc(r['뜻'])}</b> {esc(r['음'])}</div>"
            f"<div class='g'>정답률 {r['정답률(%)']}%</div></div>"
            for _, r in weak.iterrows()) + "</div>")

    if s.words:
        st.markdown("##### 복습 단계별 한자 수")
        counts = [0] * len(INTERVALS)
        unseen = 0
        for w in s.words:
            e = s.srs.get(w["hanja"])
            if e is None:
                unseen += 1
            else:
                counts[min(e.get("box", 0), len(INTERVALS) - 1)] += 1
        bd = pd.DataFrame({"한자 수": [unseen] + counts}, index=["아직 안 풀어봄"] + BOX_LABELS)
        st.bar_chart(bd, color="#2B3A67", sort=False)
        st.caption("틀리면 '바로 복습'으로 돌아가고, 맞힐수록 다음 복습까지의 간격이 길어져요.")

    with st.expander("전체 표 보기"):
        st.dataframe(df, width="stretch", hide_index=True)
    with st.expander("💾 기록 백업 (브라우저 데이터가 지워질 때를 대비)"):
        st.caption("기록은 이 브라우저에 자동 저장돼요. 다만 아이폰 사파리처럼 한동안 안 들어오면 저장 공간을 비우는 "
                   "브라우저도 있어서, 가끔 백업 파일을 받아 두면 안전해요. 복구는 왼쪽 메뉴 '진행 상황 저장 / 불러오기'에서 해요.")
        st.download_button("💾 백업 파일 받기", progress_json(), "hanja_progress.json", "application/json",
                           width="stretch", key="backup_stats")


# ----------------------------------------------------------------------
# 한자 분해 (구성 글자) — 단어장에서 한자를 누르면 어떤 글자들로 이루어졌는지 보여줘요
#   休 = 亻(人 사람 인) + 木(나무 목)
# 분해 데이터: 오픈 데이터 IDS(cjkvi-ids). 프로그램 폴더에 ids.txt(또는 data/ids.txt)를 두면 그걸 쓰고,
# 없으면 처음 한 번 내려받아요. 부수 이름(훈음)은 아래 표 + 내 단어장에 있는 한자의 뜻·음을 써요.
# ----------------------------------------------------------------------
IDS_URLS = [
    "https://raw.githubusercontent.com/cjkvi/cjkvi-ids/master/ids.txt",
    "https://cdn.jsdelivr.net/gh/cjkvi/cjkvi-ids@master/ids.txt",  # 예비 주소
]
IDS_RETRY_SEC = 120  # 내려받기에 실패하면 이 시간 동안은 다시 시도하지 않아요
BU_CARD_LIMIT = 800  # 카드가 너무 많으면 화면이 무거워져서, 앞쪽 이만큼만 눌러서 볼 수 있게 해요

# 강희 214 부수 (훈음은 한국 한자 자전 기준)
RADICAL_TEXT = """
一 한 일|丨 뚫을 곤|丶 점 주|丿 삐침 별|乙 새 을|亅 갈고리 궐|二 두 이|亠 돼지해머리 두|人 사람 인|儿 어진사람 인
入 들 입|八 여덟 팔|冂 멀 경|冖 덮을 멱|冫 이수변 빙|几 안석 궤|凵 입벌릴 감|刀 칼 도|力 힘 력|勹 쌀 포
匕 비수 비|匚 상자 방|匸 감출 혜|十 열 십|卜 점 복|卩 병부 절|厂 기슭 엄|厶 사사 사|又 또 우|口 입 구
囗 에울 위|土 흙 토|士 선비 사|夂 뒤져올 치|夊 천천히 걸을 쇠|夕 저녁 석|大 큰 대|女 여자 녀|子 아들 자|宀 갓머리 면
寸 마디 촌|小 작을 소|尢 절름발이 왕|尸 주검 시|屮 싹날 철|山 메 산|巛 개미허리 천|工 장인 공|己 몸 기|巾 수건 건
干 방패 간|幺 작을 요|广 집 엄|廴 민책받침 인|廾 스물입발 공|弋 주살 익|弓 활 궁|彐 튼가로왈 계|彡 터럭 삼|彳 두인변 척
心 마음 심|戈 창 과|戶 지게 호|手 손 수|支 지탱할 지|攴 칠 복|文 글월 문|斗 말 두|斤 도끼 근|方 모 방
无 없을 무|日 날 일|曰 가로 왈|月 달 월|木 나무 목|欠 하품 흠|止 그칠 지|歹 뼈앙상할 알|殳 몽둥이 수|毋 말 무
比 견줄 비|毛 터럭 모|氏 성씨 씨|气 기운 기|水 물 수|火 불 화|爪 손톱 조|父 아비 부|爻 사귈 효|爿 조각 장
片 조각 편|牙 어금니 아|牛 소 우|犬 개 견|玄 검을 현|玉 구슬 옥|瓜 오이 과|瓦 기와 와|甘 달 감|生 날 생
用 쓸 용|田 밭 전|疋 필 필|疒 병질 엄|癶 필발머리 발|白 흰 백|皮 가죽 피|皿 그릇 명|目 눈 목|矛 창 모
矢 화살 시|石 돌 석|示 보일 시|禸 짐승발자국 유|禾 벼 화|穴 구멍 혈|立 설 립|竹 대 죽|米 쌀 미|糸 실 사
缶 장군 부|网 그물 망|羊 양 양|羽 깃 우|老 늙을 로|而 말이을 이|耒 쟁기 뢰|耳 귀 이|聿 붓 율|肉 고기 육
臣 신하 신|自 스스로 자|至 이를 지|臼 절구 구|舌 혀 설|舛 어그러질 천|舟 배 주|艮 그칠 간|色 빛 색|艸 풀 초
虍 범호엄 호|虫 벌레 충|血 피 혈|行 다닐 행|衣 옷 의|襾 덮을 아|見 볼 견|角 뿔 각|言 말씀 언|谷 골 곡
豆 콩 두|豕 돼지 시|豸 해태 치|貝 조개 패|赤 붉을 적|走 달릴 주|足 발 족|身 몸 신|車 수레 거|辛 매울 신
辰 별 진|辵 책받침 착|邑 고을 읍|酉 닭 유|釆 분별할 변|里 마을 리|金 쇠 금|長 길 장|門 문 문|阜 언덕 부
隶 미칠 이|隹 새 추|雨 비 우|靑 푸를 청|非 아닐 비|面 낯 면|革 가죽 혁|韋 다룬가죽 위|韭 부추 구|音 소리 음
頁 머리 혈|風 바람 풍|飛 날 비|食 밥 식|首 머리 수|香 향기 향|馬 말 마|骨 뼈 골|高 높을 고|髟 터럭발 표
鬥 싸울 투|鬯 울창주 창|鬲 솥 력|鬼 귀신 귀|魚 물고기 어|鳥 새 조|鹵 짠땅 로|鹿 사슴 록|麥 보리 맥|麻 삼 마
黃 누를 황|黍 기장 서|黑 검을 흑|黹 바느질할 치|黽 맹꽁이 맹|鼎 솥 정|鼓 북 고|鼠 쥐 서|鼻 코 비|齊 가지런할 제
齒 이 치|龍 용 룡|龜 거북 귀|龠 피리 약
"""
# 부수는 아니지만 자주 나오는 구성 글자
EXTRA_TEXT = """
王 임금 왕|中 가운데 중|千 일천 천|也 어조사 야|良 어질 량|可 옳을 가|才 재주 재|永 길 영|京 서울 경|少 적을 소
"""
# 모양이 바뀐 부수 → 원래 부수 (예: 亻은 人이 왼쪽에 올 때의 모양)
VARIANT = {
    "亻": "人", "氵": "水", "氺": "水", "扌": "手", "忄": "心", "㣺": "心", "⺖": "心", "犭": "犬", "艹": "艸",
    "⺾": "艸", "礻": "示", "衤": "衣", "罒": "网", "⺫": "目", "刂": "刀", "灬": "火", "爫": "爪", "牜": "牛",
    "糹": "糸", "飠": "食", "⺼": "肉", "攵": "攴", "辶": "辵", "⻌": "辵", "耂": "老", "⺮": "竹", "彑": "彐",
    "覀": "襾", "歺": "歹", "青": "靑", "戸": "戶", "户": "戶", "黄": "黃", "𧾷": "足", "⻏": "邑", "⻖": "阜",
}
# 왼쪽/오른쪽에 따라 둘로 갈리는 모양
SPECIAL = {"阝": ("언덕 부 · 고을 읍", "왼쪽에 있으면 阜(언덕), 오른쪽에 있으면 邑(고을)")}


def _kv_table(text):
    out = {}
    for item in text.replace("\n", "|").split("|"):
        item = item.strip()
        if item:
            ch, _, mean = item.partition(" ")
            out[ch] = mean.strip()
    return out


RADICALS = _kv_table(RADICAL_TEXT)
EXTRA_COMP = _kv_table(EXTRA_TEXT)

IDC_NAMES = {
    "⿰": "왼쪽·오른쪽으로 나란히", "⿱": "위·아래로 쌓음", "⿲": "왼쪽·가운데·오른쪽 세 부분",
    "⿳": "위·가운데·아래 세 부분", "⿴": "바깥 글자가 안쪽을 에워쌈", "⿵": "위에서 감싸 내려옴",
    "⿶": "아래에서 감싸 올라감", "⿷": "왼쪽에서 감싸 안음", "⿸": "왼쪽 위에서 감쌈",
    "⿹": "오른쪽 위에서 감쌈", "⿺": "왼쪽 아래에서 감쌈", "⿻": "서로 겹쳐진 모양",
    "⿼": "서로 겹쳐진 모양", "⿽": "서로 겹쳐진 모양",
}
IDC_ARITY = {c: 2 for c in "⿰⿱⿴⿵⿶⿷⿸⿹⿺⿻⿼⿽㇯"}
IDC_ARITY.update({"⿲": 3, "⿳": 3, "⿾": 1, "⿿": 1})
_IDS_TOKEN = re.compile(r"&[^;\s]+;|.", re.S)


@st.cache_resource(show_spinner=False)
def _ids_state():
    """분해 표를 화면이 다시 그려져도 기억해요 (실행마다 새로 내려받지 않게)"""
    return {"table": None, "source": "", "fail_at": 0.0}


_IDS = _ids_state()


def parse_ids_text(text):
    """ids.txt 한 줄: 'U+4F11<탭>休<탭>⿰亻木'  (한국 자형 표시 [K]가 있으면 그걸 우선)"""
    table = {}
    for ln in text.splitlines():
        if not ln or ln.startswith(";"):
            continue
        parts = ln.split("\t")
        if len(parts) < 3 or not parts[0].startswith("U+"):
            continue
        ch = parts[1]
        if len(ch) != 1 or ord(ch) >= 0x30000:  # &CDP-…; 같은 특수 항목은 건너뜀
            continue
        best = None
        for cand in parts[2:]:
            m = re.match(r"^(.*?)(?:\[([A-Z]+)\])?$", cand.strip())
            ids, tag = m.group(1), m.group(2) or ""
            if not ids:
                continue
            if best is None:
                best = ids
            if "K" in tag:
                best = ids
                break
        if best and best != ch:  # 자기 자신뿐이면 더 나눌 수 없는 글자
            table[ch] = best
    return table


def _ids_paths():
    here = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else "."
    return [os.path.join(here, "ids.txt"), os.path.join(here, "data", "ids.txt"), "ids.txt", "data/ids.txt",
            os.path.join(tempfile.gettempdir(), "hanja_ids", "ids.txt")]


def load_ids():
    """분해 표를 돌려줘요. 없으면 내려받고, 실패하면 None (한동안 다시 시도하지 않아요)"""
    if _IDS["table"]:
        return _IDS["table"]
    for p in _ids_paths():
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    t = parse_ids_text(f.read())
                if len(t) > 1000:
                    _IDS.update(table=t, source=p)
                    return t
            except Exception:
                continue
    if time.time() - _IDS["fail_at"] < IDS_RETRY_SEC:
        return None
    cache = os.path.join(tempfile.gettempdir(), "hanja_ids", "ids.txt")
    for url in IDS_URLS:
        try:
            with urllib.request.urlopen(url, timeout=30) as r:
                data = r.read()
            t = parse_ids_text(data.decode("utf-8"))
            if len(t) > 1000:
                try:
                    os.makedirs(os.path.dirname(cache), exist_ok=True)
                    with open(cache + ".part", "wb") as f:
                        f.write(data)
                    os.replace(cache + ".part", cache)
                except Exception:
                    pass
                _IDS.update(table=t, source=url)
                return t
        except Exception:
            continue
    _IDS["fail_at"] = time.time()
    return None


def load_ids_with_notice():
    if _IDS["table"]:
        return _IDS["table"]
    with st.spinner("한자 구성 데이터를 불러오는 중이에요… (처음 한 번만, 몇 초 걸려요)"):
        t = load_ids()
    if not t:
        st.warning("한자 구성 데이터를 불러오지 못했어요. 인터넷 연결을 확인하고 잠시 뒤 다시 켜 주세요. "
                   "(또는 cjkvi-ids의 ids.txt 파일을 hanja.py와 같은 폴더에 넣어 두면 바로 써요)")
    return t


def parse_ids_tree(ids):
    """앞에서부터 읽는 IDS('⿰亻木')를 트리로: 글자 또는 (구조기호, [자식…])"""
    toks = _IDS_TOKEN.findall(ids)
    pos = 0

    def node():
        nonlocal pos
        if pos >= len(toks):
            return "?"
        t = toks[pos]
        pos += 1
        n = IDC_ARITY.get(t)
        if n:
            kids = []
            for _ in range(n):
                kids.append(node())
            return (t, kids)
        return t

    return node()


def _leaves(node):
    if isinstance(node, tuple):
        out = []
        for k in node[1]:
            out += _leaves(k)
        return out
    return [node]


def comp_info(c, by):
    """구성 글자 하나의 설명: 훈음, 부수 여부, 변형 설명"""
    if c in SPECIAL:
        return {"mean": SPECIAL[c][0], "radical": True, "note": SPECIAL[c][1]}
    base = VARIANT.get(c, c)
    w = by.get(c) or by.get(base)
    mean = f"{w['meaning']} {w['sound']}".strip() if w and (w.get("meaning") or w.get("sound")) else ""
    if not mean:
        mean = RADICALS.get(base) or EXTRA_COMP.get(base) or ""
    note = f"{base}({RADICALS.get(base, '').strip() or '…'})이 모양을 바꾼 글자" if base != c else ""
    return {"mean": mean, "radical": base in RADICALS, "note": note}


def decomp_tree(ch, table, depth=2):
    """한자 하나를 구성 글자로 쪼개요. 더 못 쪼개면 None.
    {'idc':구조기호, 'parts':[{'ch':글자, 'n':반복 횟수, 'sub':하위 분해 또는 None}]}"""
    ids = table.get(ch)
    if not ids:
        return None
    tree = parse_ids_tree(ids)
    if not isinstance(tree, tuple):
        return None
    parts = []
    for leaf in _leaves(tree):
        if parts and parts[-1]["ch"] == leaf:
            parts[-1]["n"] += 1
            continue
        parts.append({"ch": leaf, "n": 1, "sub": None})
    if depth > 1:
        for p in parts:
            c = p["ch"]
            if (len(c) == 1 and c not in VARIANT and c not in SPECIAL and c not in EXTRA_COMP
                    and VARIANT.get(c, c) not in RADICALS):
                p["sub"] = decomp_tree(c, table, depth - 1)
    return {"idc": tree[0], "parts": parts}


def _part_html(p, by):
    c = p["ch"]
    if len(c) > 1:  # &CDP-8BF4; 처럼 화면에 그릴 수 없는 조각
        return "<span class='bu-part'><b>▯</b><i>글자 없음</i></span>"
    info = comp_info(c, by)
    cls = "bu-part rad" if info["radical"] else "bu-part"
    cnt = f"<em>×{p['n']}</em>" if p["n"] > 1 else ""
    title = f" title='{esc(info['note'])}'" if info["note"] else ""
    return f"<span class='{cls}'{title}><b>{esc(c)}</b><i>{esc(info['mean'] or '뜻 정보 없음')}</i>{cnt}</span>"


def _eq_html(ch, tr, by):
    parts = "<span class='bu-op'>+</span>".join(_part_html(p, by) for p in tr["parts"])
    return f"<div class='bu-eq'><span class='bu-big'>{esc(ch)}</span><span class='bu-op'>=</span>{parts}</div>"


def _sub_html(tr, by):
    out = ""
    for p in tr["parts"]:
        # 아는 글자가 하나도 없이 획 조각만 나오는 경우는 보여 줘도 도움이 안 돼서 숨겨요
        if p["sub"] and any(len(q["ch"]) == 1 and (comp_info(q["ch"], by)["mean"]) for q in p["sub"]["parts"]):
            out += (f"<div class='bu-sub'>{_eq_html(p['ch'], p['sub'], by)}"
                    f"<div class='bu-note'>{esc(IDC_NAMES.get(p['sub']['idc'], ''))}</div></div>")
    return out


def decomp_html(word, by, table):
    """단어(한 글자 이상)의 글자마다 구성 글자를 보여주는 HTML 조각"""
    blocks = []
    for ch in list(dict.fromkeys(c for c in str(word) if has_hanja(c)))[:6]:
        tr = decomp_tree(ch, table)
        if tr:
            blocks.append(f"<div class='bu-char'>{_eq_html(ch, tr, by)}"
                          f"<div class='bu-note'>{esc(IDC_NAMES.get(tr['idc'], ''))}</div>{_sub_html(tr, by)}</div>")
            continue
        info = comp_info(ch, by)
        if ch in table:
            msg = "분해 정보를 찾지 못했어요."
        elif info["radical"]:
            msg = f"더 나눌 수 없는 기본 글자예요 · <b>{esc(info['mean'])}</b> (부수)"
        else:
            msg = "더 나눌 수 없는 기본 글자예요."
            if info["mean"]:
                msg += f" · <b>{esc(info['mean'])}</b>"
        blocks.append(f"<div class='bu-char'><div class='bu-eq'><span class='bu-big'>{esc(ch)}</span>"
                      f"<span class='bu-note'>{msg}</span></div></div>")
    if not blocks:
        return "<div class='bu-note'>한자가 아니라 분해할 수 없어요.</div>"
    return "".join(blocks)


def _terms(q):
    return [t for t in re.split(r"\s+", str(q).strip().lower()) if t]


def word_matches(w, terms):
    """검색어(띄어쓰기로 여러 개 가능)가 모두 한자·뜻·음·분류 중 어딘가에 들어 있으면 통과.
    필드를 이어 붙이지 않고 따로 비교해서, '하늘'+'천' 사이에 걸친 '늘천' 같은 엉뚱한 결과가 안 나와요"""
    fields = [w["hanja"].lower(), w["meaning"].lower(), w["sound"].lower(), w.get("group", "").lower()]
    return all(any(t in f for f in fields) for t in terms)


def highlight(text, terms):
    """검색어와 겹치는 부분에 형광펜 (HTML 안전)"""
    text = str(text)
    if not terms:
        return esc(text)
    pat = re.compile("(" + "|".join(re.escape(t) for t in sorted(terms, key=len, reverse=True)) + ")", re.I)
    return "".join(f"<mark>{esc(p)}</mark>" if i % 2 else esc(p) for i, p in enumerate(pat.split(text)))


def clear_study_search():
    st.session_state.study_q = ""


def render_study():
    s = st.session_state
    show_notice("word", "study_top")
    with st.form("study_form", border=False):
        c1, c2 = st.columns([5, 2], vertical_alignment="bottom")
        q = c1.text_input("한자·뜻·음 검색", key="study_q", placeholder="🔍 예: 天 / 하늘 / 천  (여러 개는 띄어쓰기)")
        c2.form_submit_button("🔍 검색", type="primary", width="stretch")
    terms = _terms(q)
    c1, c2 = st.columns(2)
    hide = c1.toggle("암기 모드 (뜻·음 가리기, 눌러서 확인)", value=False)
    view = c2.radio("보기", ["카드", "표"], horizontal=True, label_visibility="collapsed")
    bu_on = st.toggle("🧩 한자 눌러서 구성 글자(부수) 보기", value=False, key="st_bu",
                      help="카드에서는 한자를 누르면 펼쳐지고, 표에서는 줄을 고르면 아래에 나와요.")
    ids_table = load_ids_with_notice() if bu_on else None
    bu = bool(ids_table)
    by_h = {w["hanja"]: w for w in s.words} if bu else {}
    flt = st.radio("모아 보기", ["전체", "★ 별표", "📝 메모"], horizontal=True, key="st_filter",
                   label_visibility="collapsed")
    total = len(s.words)
    words = [w for w in s.words if word_matches(w, terms)] if terms else list(s.words)
    if flt == "★ 별표":
        words = [w for w in words if is_star(w["hanja"])]
    elif flt == "📝 메모":
        words = [w for w in words if note_of(w["hanja"])]
    if terms:  # 글자가 정확히 같은 한자·음을 앞으로
        words.sort(key=lambda w: 0 if any(t in (w["hanja"].lower(), w["sound"].lower()) for t in terms) else 1)
        r1, r2 = st.columns([5, 2], vertical_alignment="center")
        r1.caption(f"'{q.strip()}' 검색 결과 {len(words)}개 (전체 {total}개 중)")
        r2.button("✖ 검색 지우기", on_click=clear_study_search, width="stretch")
        if not words:
            st.info("검색 결과가 없어요. 한자를 입력하기 어렵다면 뜻이나 음(예: 하늘, 천)으로 찾아보세요.")
    else:
        st.caption(f"{total}개 · 검색어를 입력하고 엔터(또는 검색 버튼)를 누르세요")
    render_note_tools(words)
    if view == "표":
        df = pd.DataFrame(words).rename(columns={"hanja": "한자", "meaning": "뜻", "sound": "음", "group": "분류"})
        if len(df):
            df.insert(0, "★", ["★" if is_star(h) else "" for h in df["한자"]])
            df["메모"] = [note_of(h) for h in df["한자"]]
        if hide and len(df):
            df["뜻"], df["음"] = "???", "?"
        if bu and len(df):
            ev = st.dataframe(df, width="stretch", hide_index=True, height=480, on_select="rerun",
                              selection_mode="single-row", key="st_tbl_sel")
            rows = list(ev.selection.rows) if ev is not None and ev.selection is not None else []
            if rows and rows[0] < len(words):
                html(f"<div class='bu-body open'>{decomp_html(words[rows[0]]['hanja'], by_h, ids_table)}</div>")
            else:
                st.caption("🧩 표에서 줄 맨 왼쪽 칸을 누르면 아래에 그 한자의 구성 글자가 나와요.")
        else:
            st.dataframe(df, width="stretch", hide_index=True, height=480)
    else:
        cards = []
        if bu and len(words) > BU_CARD_LIMIT:
            st.caption(f"한자가 많아서 앞쪽 {BU_CARD_LIMIT}개만 눌러서 볼 수 있어요. 검색으로 좁히면 모두 볼 수 있어요.")
        for i, w in enumerate(words):
            mean = f"<b>{highlight(w['meaning'], terms)}</b> {highlight(w['sound'], terms)}"
            body = f"<details><summary>보기</summary>{mean}</details>" if hide else mean
            g = f"<div class='g'>{esc(w['group'])}</div>" if w.get("group") else ""
            star = "<span class='star'>★</span>" if is_star(w["hanja"]) else ""
            nt = note_of(w["hanja"])
            note_html = f"<div class='note' title='{esc(nt)}'>📝 {esc(nt)}</div>" if nt else ""
            if bu and i < BU_CARD_LIMIT:  # 한자를 누르면 펼쳐지는 구성 글자 (화면을 새로 그리지 않아요)
                head = f"<details class='bu'><summary>{highlight(w['hanja'], terms)}</summary></details>"
                tail = f"<div class='bu-body'>{decomp_html(w['hanja'], by_h, ids_table)}</div>"
            else:
                head, tail = f"<div class='h'>{highlight(w['hanja'], terms)}</div>", ""
            cards.append(f"<div class='wcard'>{star}{head}<div class='m'>{body}</div>{g}{note_html}{tail}</div>")
        html("<div class='wgrid'>" + "".join(cards) + "</div>")
    st.divider()
    render_set_tools()
    render_sheet()


# --- 시험지(PDF) ----------------------------------------------------------------
def _dashed(d, x0, y0, x1, y1, fill, dash=8):
    if x0 == x1:
        for t in range(int(y0), int(y1), dash * 2):
            d.line([(x0, t), (x0, min(t + dash, y1))], fill=fill, width=2)
    else:
        for t in range(int(x0), int(x1), dash * 2):
            d.line([(t, y0), (min(t + dash, x1), y0)], fill=fill, width=2)


def make_sheet_pdf(words, kind, font_path, title="한자 시험지"):
    W, H, M = 1240, 1754, 90  # A4, 150dpi
    ink, soft, red = (30, 36, 56), (110, 118, 140), (225, 175, 170)
    f_title, f_num = load_font(font_path, 54), load_font(font_path, 34)
    f_big, f_txt, f_small = load_font(font_path, 74), load_font(font_path, 44), load_font(font_path, 30)
    pages = []

    def page(first):
        im = Image.new("RGB", (W, H), "white")
        d = ImageDraw.Draw(im)
        y = M
        if first:
            d.text((M, M), title, font=f_title, fill=ink)
            d.text((W - M, M + 18), f"이름 __________   점수 ____ / {len(words)}", font=f_small, fill=soft, anchor="ra")
            d.line([(M, M + 90), (W - M, M + 90)], fill=ink, width=3)
            y = M + 120
        pages.append(im)
        return d, y

    if kind == "한자 → 뜻·음 쓰기":
        row_h = 122
        d, y0 = page(True)
        y = y0
        for i, w in enumerate(words, 1):
            if y + row_h > H - M:
                d, y = page(False)
            d.text((M, y + 60), f"{i}.", font=f_num, fill=soft, anchor="lm")
            d.text((M + 90, y + 60), w["hanja"], font=f_big, fill=ink, anchor="lm")
            d.line([(M + 380, y + 100), (W - M, y + 100)], fill=soft, width=2)
            y += row_h
    else:
        row_h, col_w = 160, (W - 2 * M) // 2
        d, y0 = page(True)
        y, col = y0, 0
        for i, w in enumerate(words, 1):
            if y + row_h > H - M:
                d, y = page(False)
                col = 0
            x = M + col * col_w
            label = f"{w['meaning']} {w['sound']}".strip()
            d.text((x, y + 70), f"{i}.", font=f_num, fill=soft, anchor="lm")
            d.text((x + 62, y + 70), label, font=f_txt, fill=ink, anchor="lm")
            n = min(len(w["hanja"]), 4)
            box = min(120, (col_w - 300) // n)
            bx = x + col_w - 20 - box * n
            for k in range(n):
                x0 = bx + k * box
                d.rectangle([x0, y + 10, x0 + box, y + 10 + box], outline=ink, width=2)
                _dashed(d, x0 + box // 2, y + 10, x0 + box // 2, y + 10 + box, red)
                _dashed(d, x0, y + 10 + box // 2, x0 + box, y + 10 + box // 2, red)
            col += 1
            if col == 2:
                col, y = 0, y + row_h
        # 마지막 줄이 한 칸만 찼다면 아래 정답지가 이어서 시작
    # 정답지
    d, y = page(False)
    d.text((M, M), f"{title} - 정답", font=f_title, fill=ink)
    y = M + 120
    d.line([(M, M + 90), (W - M, M + 90)], fill=ink, width=3)
    line_h, col_w = 62, (W - 2 * M) // 2
    col = 0
    for i, w in enumerate(words, 1):
        if y + line_h > H - M:
            if col == 0:
                col, y = 1, M + 120
            else:
                d, y = page(False)
                col, y = 0, M
        x = M + col * col_w
        d.text((x, y), f"{i}. {w['hanja']}  {w['meaning']} {w['sound']}", font=f_txt, fill=ink)
        y += line_h
    buf = io.BytesIO()
    pages[0].save(buf, "PDF", save_all=True, append_images=pages[1:], resolution=150)
    return buf.getvalue()


def render_sheet():
    s = st.session_state
    with st.expander("🖨️ 시험지 만들기 (PDF)"):
        font = find_font()
        if not font:
            st.info("한자 폰트가 없어 시험지를 만들 수 없어요.")
            return
        kind = st.radio("시험 유형", ["한자 → 뜻·음 쓰기", "뜻·음 → 한자 쓰기"], horizontal=True, key="sheet_kind")
        n = st.slider("문제 수", 5, min(60, max(5, len(s.words))), min(20, max(5, len(s.words))), step=5, key="sheet_n") \
            if len(s.words) > 5 else len(s.words)
        if st.button("시험지 만들기", key="make_sheet"):
            pool = random.sample(s.words, min(n, len(s.words)))
            with st.spinner("만드는 중..."):
                s.sheet = make_sheet_pdf(pool, kind, font)
        if s.sheet:
            st.download_button("📄 PDF 받기 (마지막 장이 정답지)", s.sheet, "한자시험지.pdf", "application/pdf",
                               type="primary", width="stretch")


def apply_preset():
    name = st.session_state.get("set_preset")
    if name in GRADE_PRESETS:
        st.session_state.update(GRADE_PRESETS[name])


def mark_custom():
    st.session_state.set_preset = "직접 설정"


def grading_sidebar():
    s = st.session_state
    with st.sidebar.expander("🎯 손글씨 채점 세부 설정"):
        st.selectbox("프리셋", ["관대", "보통", "엄격", "직접 설정"], key="set_preset", on_change=apply_preset,
                     help="한 번에 여러 설정을 바꿔요. 아래 값을 직접 건드리면 '직접 설정'이 돼요.")
        st.markdown("**기본 채점**")
        st.slider("통과 점수 (높을수록 엄격)", 0.30, 0.95, step=0.01, key="set_thr", on_change=mark_custom,
                  help="내 글씨와 정답 글자가 얼마나 겹치는지 0~1점으로 계산해서, 이 점수 이상이면 정답이에요.")
        st.slider("위치 허용 오차 (글자 크기의 %)", 1.0, 10.0, step=0.5, key="set_tol", on_change=mark_custom,
                  help="획이 정답 위치에서 이만큼 벗어나도 완전히 맞는 걸로 쳐요. 작을수록 엄격해요.")
        st.slider("위치·크기·기울기 보정", 0.0, 3.0, step=0.5, key="set_align", on_change=mark_custom,
                  help="글씨가 한쪽으로 치우치거나, 납작/길쭉하거나, 기울어도 정답에 가장 잘 맞게 자동으로 맞춰서 비교해요. 0이면 끔.")
        st.slider("빠진 획 ↔ 덧쓴 획", 0.5, 2.0, step=0.1, key="set_beta", on_change=mark_custom,
                  help="1.0은 균형. 1보다 크면 획이 빠진 것을 더 엄하게, 1보다 작으면 쓸데없는 획이 있는 것을 더 엄하게 봐요.")
        st.markdown("**획 단위 검사**")
        st.checkbox("빠진 획 / 덧쓴 획 검사", key="set_chk", on_change=mark_custom,
                    help="전체 점수가 높아도, 획 하나가 통째로 빠지거나 더 들어갔으면 오답으로 처리해요.")
        st.slider("허용 길이 (글자 크기의 %)", 4.0, 30.0, step=1.0, key="set_miss", on_change=mark_custom,
                  disabled=not s.get("set_chk", True),
                  help="이 길이 이상으로 이어서 어긋난 부분이 있으면 오답이에요. 작을수록 엄격하고, 점·짧은 획까지 따져요.")
        st.markdown("**비슷한 글자 구분**")
        st.checkbox("비슷한 다른 글자와 비교", key="set_disc", on_change=mark_custom,
                    help="단어장 안의 모양이 닮은 글자(예: 大·太·犬)와 비교해서, 그쪽이 더 비슷하면 오답으로 처리해요.")
        st.slider("점수 차 허용", 0.00, 0.20, step=0.01, key="set_margin", on_change=mark_custom,
                  disabled=not s.get("set_disc", True),
                  help="다른 글자가 정답보다 이 점수 넘게 더 비슷할 때만 오답이에요. 클수록 관대해요.")
        st.checkbox("글꼴 모양 차이 허용 (한국·중국·일본식)", key="set_variants", on_change=mark_custom,
                    help="같은 한자도 나라별 글꼴에서 모양이 조금 달라서, 가능하면 여러 모양을 모두 정답으로 인정해요.")
        if st.button("기본값으로 되돌리기"):
            s.update(GRADE_DEFAULTS)
            s.set_preset = "보통"
            st.rerun()
        if s.calib:
            st.markdown("**💡 채점 기준 추천**")
            scores = [c["score"] for c in s.calib if c["why"] == "score"]
            others = [c for c in s.calib if c["why"] != "score"]
            st.caption(f"'내가 맞게 썼어요'로 정정한 손글씨 {len(s.calib)}건을 참고했어요.")
            if scores:
                sug = max(0.30, round(min(scores) - 0.01, 2))
                if sug < s.get("set_thr", 0.7):
                    st.button(f"통과 점수를 {sug:.2f}로 낮추기", on_click=lambda v=sug: s.update(set_thr=v, set_preset="직접 설정"))
            if any(c["why"] == "stroke" for c in others):
                st.caption("획 검사 때문에 오답 처리된 정정이 있어요 → '허용 길이'를 늘려 보세요.")
            if any(c["why"] == "disc" for c in others):
                st.caption("비슷한 글자 비교 때문에 오답 처리된 정정이 있어요 → '점수 차 허용'을 늘려 보세요.")
            st.button("추천 기록 지우기", on_click=lambda: s.calib.clear())


# ----------------------------------------------------------------------
# 사이드바 + 메인
# ----------------------------------------------------------------------
def sidebar():
    s = st.session_state
    st.sidebar.title("📖 메뉴")
    if s.words:
        with st.sidebar.expander("📂 단어장 바꾸기"):
            load_panel()
    st.sidebar.number_input("오늘의 목표 (문제 수)", 5, 500, step=5, key="set_goal")
    st.sidebar.slider("펜 굵기", 3, 20, key="set_stroke")
    st.sidebar.slider("쓰는 칸 크기", 120, 360, step=20, key="set_canvas_px",
                      help="한자를 쓰는 칸의 크기예요. 화면이 작으면 줄이세요. (바꾸면 쓰던 글씨는 지워져요)")
    grading_sidebar()

    with st.sidebar.expander("💾 진행 상황 저장 / 불러오기"):
        st.caption("평소엔 이 기기에 자동 저장돼요. 다른 기기로 옮기거나 백업할 때만 쓰세요. (오답 노트·통계·복습 기록·날짜별 기록·메모·사자성어 기록·플랜 설정)")
        st.download_button("💾 저장", progress_json(), "hanja_progress.json", "application/json", width="stretch",
                           key="backup_side")
        pf = st.file_uploader("불러오기", type=["json"], key=f"progress_file_{s.up_gen}")
        if pf is not None and (pf.name, pf.size) != s.progress_sig:
            try:
                data = json.loads(pf.getvalue().decode("utf-8"))
                merge_progress(sanitize_progress(data))
                s.progress_sig = (pf.name, pf.size)
                s.flash = "진행 상황을 불러왔어요."
            except Exception:
                s.flash = "진행 파일을 읽지 못했어요."

    with st.sidebar.expander("🧰 진단 정보 / 초기화"):
        if not find_font():
            st.warning("한자 폰트를 찾지 못하고 자동으로 내려받지도 못해서 손글씨는 자가 채점으로 진행돼요. (packages.txt에 fonts-noto-cjk가 있는지 확인)")
        if not HAS_CANVAS:
            st.error(
                "⚠️ 한자 쓰기 기능(streamlit-drawable-canvas)을 불러오지 못했어요. "
                "저장소의 requirements.txt(소문자)에 streamlit-drawable-canvas==0.9.3 이 있는지 확인하고 앱을 Reboot 해 주세요."
            )
            st.code(CANVAS_ERR)
        st.write(f"Streamlit 버전: {st.__version__}")
        st.write(f"캔버스 라이브러리: {'정상' if HAS_CANVAS else '불러오기 실패'}")
        st.write(f"한자 폰트: {find_font() or '없음'} ({_FONT['source'] or '-'})")
        st.write("자동 저장: " + {True: "켜짐", False: "이 브라우저에서는 사용할 수 없음", None: "확인 중"}[s.store_ok])
        st.caption("화면이 이상하게 동작하면 주소 끝에 ?nostore=1 을 붙여 접속해 보세요. (자동 저장 없이 실행)")
        st.divider()
        st.button("🔄 퀴즈 초기화", on_click=reset_quiz, width="stretch")
        st.checkbox("기록 초기화 확인 (오답 노트·통계·복습 기록 삭제, 단어장·메모는 유지)", key="confirm_reset")
        st.button("🗑️ 기록만 초기화", on_click=reset_all, disabled=not s.get("confirm_reset", False), width="stretch")
        st.divider()
        st.markdown("**⚠️ 앱 완전 초기화**")
        st.caption("단어장·학습 기록·메모·설정, 그리고 이 브라우저에 자동 저장된 내용까지 전부 지우고 처음 상태로 돌아가요. "
                   "되돌릴 수 없으니, 필요하면 먼저 위의 '💾 진행 상황 저장'으로 백업을 받아 두세요.")
        st.checkbox("전부 지워도 괜찮아요", key="confirm_factory")
        st.button("☢️ 앱 완전 초기화", on_click=factory_reset, type="primary",
                  disabled=not s.get("confirm_factory", False), width="stretch")


def run_app():
    sidebar()
    refresh_idioms()  # 단어장 한자 → 사자성어 자동 연결
    s = st.session_state
    header()
    if not s.words:
        if s.flash:
            st.toast(s.flash)
            s.flash = ""
        render_welcome()
        if s.words:  # 방금 불러왔으면 화면을 새로 그려서 탭을 보여줌
            st.rerun()
        return
    if s.flash:  # 레이아웃을 밀지 않는 토스트로 보여줘야 탭 선택이 풀리지 않아요
        st.toast(s.flash)
        s.flash = ""
    if not HAS_CANVAS:
        st.warning("한자 쓰기 기능을 불러오지 못해 지금은 손글씨 문제를 풀 수 없어요. 왼쪽 메뉴의 '진단 정보'를 확인해 주세요.")
    t1, t2, t3, t4, t5, t6 = st.tabs(
        ["✏️ 연습", "🃏 카드", "📜 사자성어", "📕 오답", "📊 기록", "📚 단어장"])
    with t1:
        render_practice()
    with t2:
        render_cards()
    with t3:
        render_idioms()
    with t4:
        render_wrong()
    with t5:
        render_stats()
    with t6:
        render_study()


# ----------------------------------------------------------------------
# 내장 사자성어 사전 — 단어장 한자로만 이루어진 사자성어를 자동으로 골라 써요
# ----------------------------------------------------------------------
IDIOM_DB_TEXT = """
街談巷說|가담항설|거리와 항간에 떠도는 소문이나 이야기
刻骨難忘|각골난망|입은 은혜가 뼈에 새겨져 잊히지 않음
刻舟求劍|각주구검|어리석고 융통성이 없어 세상 변화를 모름
肝膽相照|간담상조|서로 마음을 터놓고 진심으로 사귐
甘呑苦吐|감탄고토|비위에 맞으면 취하고 맞지 않으면 버림
甘言利說|감언이설|남의 비위를 맞추는 달콤한 말과 이로운 조건으로 꾀는 말
感之德之|감지덕지|몹시 고맙게 여김
改過遷善|개과천선|지난 잘못을 고치고 착하게 됨
去頭截尾|거두절미|머리와 꼬리를 자름: 앞뒤 다 빼고 요점만 말함
居安思危|거안사위|편안할 때에도 닥칠 위험을 생각하고 대비함
去者日疎|거자일소|떠난 사람은 날이 갈수록 멀어짐
乾坤一擲|건곤일척|운명을 걸고 단판으로 승패를 겨룸
牽强附會|견강부회|이치에 맞지 않는 말을 억지로 끌어다 붙임
隔靴搔癢|격화소양|신 신고 발바닥 긁기: 일이 성에 차지 않아 답답함
見蚊拔劍|견문발검|모기를 보고 칼을 뽑음: 사소한 일에 크게 성냄
見物生心|견물생심|물건을 보면 갖고 싶은 마음이 생김
犬馬之勞|견마지로|윗사람에게 바치는 자신의 노력을 낮추어 이르는 말
犬猿之間|견원지간|개와 원숭이 사이: 몹시 사이가 나쁜 관계
見危授命|견위수명|나라가 위태로울 때 목숨을 바침
結草報恩|결초보은|죽어서도 은혜를 잊지 않고 갚음
結者解之|결자해지|일을 저지른 사람이 그 일을 풀어야 함
謙讓之德|겸양지덕|겸손하고 사양하는 미덕
傾國之色|경국지색|나라를 기울게 할 만큼 뛰어난 미인
敬而遠之|경이원지|겉으로는 공경하되 가까이하지는 않음
驚天動地|경천동지|세상을 놀라게 함
經世濟民|경세제민|세상을 다스려 백성을 구제함
鷄卵有骨|계란유골|운이 나쁜 사람은 하는 일마다 안 풀림
鷄鳴狗盜|계명구도|하찮은 재주도 때로는 쓸모가 있음
孤軍奮鬪|고군분투|도움 없이 홀로 힘겹게 애씀
高臺廣室|고대광실|매우 크고 좋은 집
孤立無援|고립무원|고립되어 도움을 받을 곳이 없음
孤掌難鳴|고장난명|손바닥 하나로는 소리를 내기 어려움: 혼자서는 일을 이루기 어려움
鼓腹擊壤|고복격양|배를 두드리고 땅을 침: 태평한 세월을 즐김
苦盡甘來|고진감래|고생 끝에 즐거움이 옴
曲學阿世|곡학아세|학문을 굽혀 세상에 아첨함
骨肉相爭|골육상쟁|가까운 혈족끼리 서로 다툼
骨肉之親|골육지친|부모·형제·자식 등 가까운 혈육
空中樓閣|공중누각|근거 없는 가공의 사물이나 이론
公明正大|공명정대|공평하고 올바르며 사사로움이 없음
公平無私|공평무사|공평하여 사사로움이 없음
過猶不及|과유불급|지나친 것은 미치지 못한 것과 같음
管鮑之交|관포지교|우정이 아주 두터운 사귐
刮目相對|괄목상대|눈을 비비고 다시 볼 만큼 실력이 크게 늘었음
矯角殺牛|교각살우|작은 흠을 고치려다 큰일을 그르침
巧言令色|교언영색|아첨하는 말과 알랑거리는 얼굴빛
敎學相長|교학상장|가르치고 배우는 가운데 서로 성장함
九曲肝腸|구곡간장|굽이굽이 서린 마음속
口蜜腹劍|구밀복검|입에는 꿀이 있고 배에는 칼이 있음: 겉으로는 친절하나 속은 음흉함
九死一生|구사일생|여러 번 죽을 고비를 넘기고 겨우 살아남
口尙乳臭|구상유취|입에서 젖내가 남: 말과 행동이 유치함
九牛一毛|구우일모|많은 것 가운데 극히 적은 일부
口是心非|구시심비|입으로는 옳다 하나 마음은 그르다 함
九折羊腸|구절양장|아홉 번 꺾인 양의 창자: 꼬불꼬불한 길
群鷄一鶴|군계일학|많은 닭 가운데 한 마리 학: 여럿 중 뛰어난 사람
群雄割據|군웅할거|여러 영웅이 각지에서 세력을 다툼
君臣有義|군신유의|임금과 신하 사이에는 의리가 있어야 함
權謀術數|권모술수|목적을 위해 온갖 수단과 방법을 꾸밈
勸善懲惡|권선징악|착한 일을 권하고 악한 일을 벌함
捲土重來|권토중래|한 번 실패한 뒤 힘을 길러 다시 쳐들어옴
金科玉條|금과옥조|소중히 여겨 꼭 지켜야 할 법칙이나 규정
金蘭之契|금란지계|친구 사이의 매우 두터운 정
金石之交|금석지교|쇠와 돌처럼 변치 않는 굳은 사귐
今昔之感|금석지감|지금과 옛날을 비교할 때 느끼는 큰 차이감
金城鐵壁|금성철벽|매우 튼튼하고 견고한 것
錦上添花|금상첨화|좋은 일 위에 또 좋은 일이 더해짐
錦衣還鄕|금의환향|비단옷을 입고 고향에 돌아옴: 성공하여 돌아옴
金枝玉葉|금지옥엽|임금의 가족, 귀한 자손
氣高萬丈|기고만장|일이 뜻대로 잘되어 매우 우쭐함
奇巖怪石|기암괴석|기이하게 생긴 바위와 돌
起死回生|기사회생|거의 죽을 뻔하다가 다시 살아남
氣盡脈盡|기진맥진|기운과 힘이 다 빠져 없어짐
騎虎之勢|기호지세|이미 내친 일이라 중도에 그만둘 수 없는 형세
吉凶禍福|길흉화복|좋은 일과 나쁜 일, 재앙과 복
落花流水|낙화유수|떨어지는 꽃과 흐르는 물: 봄의 끝 경치, 쇠퇴하는 모양
難兄難弟|난형난제|누구를 형이라 하고 아우라 하기 어려움: 우열을 가리기 어려움
南柯一夢|남가일몽|덧없는 한때의 부귀영화
南橘北枳|남귤북지|환경에 따라 사람의 성질이 변함
男女有別|남녀유별|남자와 여자 사이에는 분별이 있어야 함
男負女戴|남부여대|남자는 지고 여자는 임: 가난한 사람들이 떠돌아다님
囊中之錐|낭중지추|주머니 속의 송곳: 재능이 뛰어난 사람은 저절로 드러남
內憂外患|내우외환|나라 안팎의 근심과 걱정
勞心焦思|노심초사|몹시 애쓰며 속을 태움
綠陰芳草|녹음방초|푸른 나무 그늘과 꽃다운 풀
論功行賞|논공행상|공을 따져 상을 줌
累卵之危|누란지위|쌓아 놓은 알처럼 몹시 위태로운 형세
多多益善|다다익선|많을수록 더욱 좋음
多岐亡羊|다기망양|갈림길이 많아 양을 잃음: 학문의 길이 많아 진리를 찾기 어려움
單刀直入|단도직입|군말 없이 바로 요점으로 들어감
斷金之交|단금지교|쇠도 자를 만큼 두터운 우정
丹脣皓齒|단순호치|붉은 입술과 하얀 이: 아름다운 여자
大器晩成|대기만성|큰 인물은 늦게 이루어짐
大驚失色|대경실색|몹시 놀라 얼굴빛이 하얗게 변함
大同小異|대동소이|거의 같고 조금 다름
大書特筆|대서특필|특별히 크게 써서 드러냄
大義名分|대의명분|마땅히 지켜야 할 도리와 본분
同價紅裳|동가홍상|같은 값이면 다홍치마
同苦同樂|동고동락|괴로움과 즐거움을 함께함
東問西答|동문서답|묻는 말에 엉뚱한 대답을 함
同病相憐|동병상련|같은 처지의 사람끼리 서로 가엾게 여김
同床異夢|동상이몽|같은 자리에 있어도 속으로는 딴생각을 함
東西古今|동서고금|동양과 서양, 옛날과 지금: 언제 어디서나
同姓同本|동성동본|성도 본관도 같음
東奔西走|동분서주|사방으로 이리저리 바쁘게 돌아다님
同族相殘|동족상잔|같은 민족끼리 서로 해치며 싸움
登高自卑|등고자비|높은 곳에 오르려면 낮은 곳에서부터 시작해야 함
燈下不明|등하불명|등잔 밑이 어두움: 가까운 일을 오히려 잘 모름
馬耳東風|마이동풍|남의 말을 귀담아듣지 않고 흘려버림
莫上莫下|막상막하|우열을 가리기 어려움
莫逆之友|막역지우|허물없이 아주 친한 벗
萬古風霜|만고풍상|오랜 세월 겪은 많은 고생
萬事亨通|만사형통|모든 일이 뜻대로 잘됨
萬事休矣|만사휴의|모든 일이 끝장났음
晩時之歎|만시지탄|때가 늦어 기회를 놓친 것을 한탄함
滿身瘡痍|만신창이|온몸이 상처투성이: 일이 엉망이 됨
亡羊補牢|망양보뢰|양을 잃고 우리를 고침: 이미 잘못된 뒤에 손을 써도 소용없음
望洋之歎|망양지탄|남의 훌륭함을 보고 자신의 부족함을 한탄함
明鏡止水|명경지수|맑은 거울과 고요한 물: 맑고 고요한 마음
明明白白|명명백백|아주 분명하고 뚜렷함
名實相符|명실상부|이름과 실제가 서로 들어맞음
明若觀火|명약관화|불을 보듯 분명함
明哲保身|명철보신|지혜롭고 사리에 밝아 몸을 잘 보전함
目不識丁|목불식정|낫 놓고 기역자도 모름: 글자를 모르는 무식함
目不忍見|목불인견|참혹하여 차마 볼 수 없음
無不通知|무불통지|모르는 것이 없이 다 통함
無所不爲|무소불위|하지 못하는 일이 없음
無爲徒食|무위도식|하는 일 없이 놀고먹음
無往不勝|무왕불승|가는 곳마다 이김
門前成市|문전성시|찾아오는 사람이 많아 문 앞이 시장 같음
物我一體|물아일체|외물과 내가 하나가 됨
美辭麗句|미사여구|아름다운 말과 글귀
美風良俗|미풍양속|아름답고 좋은 풍속
博覽强記|박람강기|여러 책을 널리 읽고 잘 기억함
博學多識|박학다식|학문이 넓고 아는 것이 많음
拔本塞源|발본색원|폐단의 근본 원인을 완전히 뽑아 없앰
傍若無人|방약무인|곁에 사람이 없는 듯 거리낌 없이 행동함
背水之陣|배수지진|물을 등지고 친 진: 죽을 각오로 일에 임함
杯盤狼藉|배반낭자|술 마시고 노는 자리가 어지러움
背恩忘德|배은망덕|은덕을 잊고 배반함
百家爭鳴|백가쟁명|많은 학자가 자유롭게 논쟁함
白骨難忘|백골난망|죽어서도 은혜를 잊을 수 없음
百年佳約|백년가약|부부가 되자는 아름다운 약속
百年大計|백년대계|먼 앞날까지 내다보는 큰 계획
百年河淸|백년하청|아무리 기다려도 이루어지기 어려움
百發百中|백발백중|쏘는 대로 다 맞음: 계획이 모두 들어맞음
白面書生|백면서생|오직 글만 읽어 세상일에 서투른 젊은이
百戰百勝|백전백승|싸울 때마다 이김
百折不屈|백절불굴|여러 번 꺾여도 굽히지 않음
伯仲之間|백중지간|서로 우열을 가리기 어려운 사이
父傳子傳|부전자전|아버지에서 아들로 대대로 전해짐
夫唱婦隨|부창부수|남편이 이끌고 아내가 따름
附和雷同|부화뇌동|주견 없이 남의 의견에 따라 움직임
粉骨碎身|분골쇄신|뼈가 가루가 되고 몸이 부서지도록 애씀
不俱戴天|불구대천|하늘을 함께 이고 살 수 없는 원수
不問可知|불문가지|묻지 않아도 알 수 있음
不問曲直|불문곡직|옳고 그름을 묻지 않음
不撤晝夜|불철주야|밤낮을 가리지 않고 힘씀
不恥下問|불치하문|손아랫사람에게 묻는 것을 부끄러워하지 않음
悲憤慷慨|비분강개|슬프고 분하여 의분이 북받침
非一非再|비일비재|한두 번이 아님
氷炭之間|빙탄지간|얼음과 숯불 사이: 서로 맞지 않는 사이
四面楚歌|사면초가|사방이 적에게 둘러싸여 고립됨
四面春風|사면춘풍|누구에게나 호의적으로 대하는 처세
四分五裂|사분오열|여러 갈래로 갈라지고 흩어짐
士農工商|사농공상|선비·농부·장인·상인: 옛 신분 계급
事必歸正|사필귀정|모든 일은 반드시 바른 길로 돌아감
士氣衝天|사기충천|사기가 하늘을 찌를 듯 높음
四通八達|사통팔달|길이 사방팔방으로 통함
山高水長|산고수장|산은 높고 물은 길게 흐름: 군자의 덕이 높고 오래감
山川草木|산천초목|산과 내와 풀과 나무: 자연
山戰水戰|산전수전|세상의 온갖 고생과 어려움을 겪음
山紫水明|산자수명|산빛은 자주빛, 물은 맑음: 자연 경치가 아름다움
山海珍味|산해진미|산과 바다의 온갖 귀한 음식
殺身成仁|살신성인|자신을 희생하여 인(仁)을 이룸
三顧草廬|삼고초려|인재를 얻으려고 정성을 다해 노력함
三三五五|삼삼오오|서넛 또는 대여섯 사람씩 무리를 지은 모양
三旬九食|삼순구식|삼십 일에 아홉 끼니를 먹음: 몹시 가난함
三人成虎|삼인성호|거짓말도 여러 사람이 하면 믿게 됨
上善若水|상선약수|가장 좋은 것은 물과 같음
桑田碧海|상전벽해|뽕나무밭이 푸른 바다가 됨: 세상이 크게 변함
塞翁之馬|새옹지마|인생의 길흉화복은 예측하기 어려움
生老病死|생로병사|사람이 겪는 네 가지 고통: 나고 늙고 병들고 죽음
生面不知|생면부지|전혀 만난 적이 없어 알지 못하는 사람
生死苦樂|생사고락|삶과 죽음, 괴로움과 즐거움
先公後私|선공후사|공적인 일을 먼저 하고 사사로운 일은 뒤로 함
先見之明|선견지명|앞일을 미리 내다보는 지혜
雪上加霜|설상가상|눈 위에 서리가 내림: 어려운 일이 겹침
說往說來|설왕설래|서로 말이 오고 감
纖纖玉手|섬섬옥수|가냘프고 고운 여자의 손
袖手傍觀|수수방관|팔짱을 끼고 보고만 있음
水魚之交|수어지교|물과 물고기처럼 떨어질 수 없는 사이
水落石出|수락석출|물이 빠지니 돌이 드러남: 진상이 밝혀짐
手不釋卷|수불석권|손에서 책을 놓지 않음: 열심히 공부함
首丘初心|수구초심|고향을 그리워하는 마음
脣亡齒寒|순망치한|입술이 없으면 이가 시림: 서로 돕는 사이의 한쪽이 망하면 다른 쪽도 위태로움
始終一貫|시종일관|처음부터 끝까지 한결같음
始終如一|시종여일|처음과 끝이 한결같음
時機尙早|시기상조|아직 때가 이르지 않음
是是非非|시시비비|옳고 그름을 따짐
識字憂患|식자우환|아는 것이 오히려 근심거리가 됨
身土不二|신토불이|몸과 땅은 둘이 아님: 우리 땅에서 난 것이 몸에 맞음
信賞必罰|신상필벌|상과 벌을 분명하게 내림
十人十色|십인십색|사람마다 취향이 다름
十中八九|십중팔구|열 가운데 여덟이나 아홉: 거의 대부분
十匙一飯|십시일반|열 사람이 한 숟가락씩 모으면 한 사람이 먹을 밥이 됨
深思熟考|심사숙고|깊이 생각하고 곰곰이 헤아림
我田引水|아전인수|자기 논에 물 대기: 자기에게 이롭게만 생각하고 행동함
安貧樂道|안빈낙도|가난해도 편안한 마음으로 도를 즐김
眼下無人|안하무인|눈 아래 사람이 없음: 교만하여 남을 업신여김
暗中摸索|암중모색|어둠 속에서 더듬어 찾음: 단서 없이 짐작으로 찾음
弱肉强食|약육강식|약한 것이 강한 것에게 먹힘
羊頭狗肉|양두구육|겉은 훌륭하나 속은 보잘것없음
良藥苦口|양약고구|좋은 약은 입에 씀: 충고는 듣기 싫으나 이로움
羊質虎皮|양질호피|겉모양은 훌륭하나 실속이 없음
漁父之利|어부지리|둘이 다투는 사이에 제삼자가 이익을 봄
言語道斷|언어도단|말문이 막힐 만큼 어이없음
言行一致|언행일치|말과 행동이 같음
言中有骨|언중유골|예사로운 말 속에 깊은 뜻이 있음
如履薄氷|여리박빙|얇은 얼음을 밟는 것처럼 몹시 위험함
易地思之|역지사지|처지를 바꾸어 생각함
緣木求魚|연목구어|나무에 올라 물고기를 구함: 불가능한 일을 하려 함
五里霧中|오리무중|짙은 안개 속: 갈피를 잡을 수 없음
烏飛梨落|오비이락|우연히 일이 겹쳐 의심을 받음
吾鼻三尺|오비삼척|내 코가 석 자: 내 사정이 급해 남을 돌볼 겨를이 없음
吳越同舟|오월동주|원수끼리도 위기에서는 서로 돕게 됨
傲霜孤節|오상고절|서리에도 굽히지 않는 외로운 절개
烏合之卒|오합지졸|질서 없이 모인 무리
溫故知新|온고지신|옛것을 익히고 새것을 앎
臥薪嘗膽|와신상담|원수를 갚으려고 고난을 참고 견딤
外柔內剛|외유내강|겉은 부드럽고 속은 단단함
樂山樂水|요산요수|산과 물을 좋아함
龍頭蛇尾|용두사미|시작은 거창하나 끝은 보잘것없음
龍蛇飛騰|용사비등|살아 움직이는 듯한 힘찬 필치
牛耳讀經|우이독경|소 귀에 경 읽기: 말해도 알아듣지 못함
優柔不斷|우유부단|어물어물 결단력이 없음
雨後竹筍|우후죽순|비 온 뒤 죽순처럼 많이 생겨남
遠交近攻|원교근공|먼 나라와 사귀고 가까운 나라를 침
遠禍召福|원화소복|화를 물리치고 복을 불러들임
月下氷人|월하빙인|혼인을 중매하는 사람
類類相從|유유상종|같은 무리끼리 서로 사귐
有口無言|유구무언|입이 있어도 말할 수 없음: 변명할 말이 없음
有名無實|유명무실|이름만 있고 실속이 없음
有備無患|유비무환|미리 준비하면 걱정이 없음
唯我獨尊|유아독존|세상에서 내가 가장 존귀함
悠悠自適|유유자적|속세를 떠나 마음 편히 삶
隱忍自重|은인자중|마음속에 감추고 참으며 몸가짐을 신중히 함
意氣揚揚|의기양양|뜻을 이루어 기운차고 자랑스러운 모양
意氣投合|의기투합|서로 마음이 맞음
意氣消沈|의기소침|기운이 없어지고 풀이 죽음
異口同聲|이구동성|여러 사람의 말이 한결같음
以卵擊石|이란격석|계란으로 바위 치기: 도저히 이길 수 없음
理路整然|이로정연|이야기의 갈피가 질서 있게 정돈됨
以實直告|이실직고|사실 그대로 고함
以心傳心|이심전심|말하지 않아도 마음으로 서로 통함
二律背反|이율배반|서로 모순되는 두 명제가 함께 성립함
以熱治熱|이열치열|열은 열로써 다스림
利用厚生|이용후생|기구를 편리하게 쓰고 생활을 넉넉하게 함
耳目口鼻|이목구비|귀·눈·입·코: 얼굴의 생김새
因果應報|인과응보|행한 대로 결과를 받음
人面獸心|인면수심|얼굴은 사람이나 마음은 짐승 같음
人山人海|인산인해|사람이 산과 바다를 이룸: 사람이 매우 많이 모임
人事不省|인사불성|정신을 잃어 의식이 없음
人生無常|인생무상|인생이 덧없음
人之常情|인지상정|사람이면 누구나 가지는 보통 감정
一刻三秋|일각삼추|아주 짧은 시간이 삼 년같이 길게 느껴짐
一擧兩得|일거양득|한 가지 일로 두 가지 이익을 얻음
一網打盡|일망타진|한꺼번에 모조리 잡음
一問一答|일문일답|한 번 묻고 한 번 답함
一目瞭然|일목요연|한 번 보고도 분명히 알 수 있음
一石二鳥|일석이조|한 가지 일로 두 가지 이득을 얻음
一魚濁水|일어탁수|한 사람의 잘못으로 여러 사람이 피해를 입음
一言半句|일언반구|아주 짧은 말
一以貫之|일이관지|하나의 이치로 모든 것을 꿰뚫음
一日三秋|일일삼추|하루가 삼 년같이 길게 느껴짐: 몹시 애타게 기다림
一長一短|일장일단|좋은 점도 있고 나쁜 점도 있음
一場春夢|일장춘몽|한바탕의 봄꿈: 덧없는 인생
一進一退|일진일퇴|나아갔다 물러났다 함
一敗塗地|일패도지|한 번 패하여 다시 일어나지 못함
一片丹心|일편단심|변치 않는 참된 마음
一寸光陰|일촌광음|매우 짧은 시간
一觸卽發|일촉즉발|조금만 건드려도 폭발할 것 같은 위기
一心同體|일심동체|한마음 한몸: 서로 굳게 결합함
一喜一悲|일희일비|기쁨과 슬픔이 번갈아 일어남
日就月將|일취월장|나날이 달마다 발전함
日月星辰|일월성신|해와 달과 별
自家撞着|자가당착|자기 말이나 행동이 앞뒤가 맞지 않음
自强不息|자강불식|스스로 힘쓰며 쉬지 않음
自給自足|자급자족|필요한 것을 스스로 마련함
自繩自縛|자승자박|자기가 만든 줄로 자기를 묶음
自業自得|자업자득|자기가 한 일의 결과를 자기가 받음
自初至終|자초지종|처음부터 끝까지의 과정
自畵自讚|자화자찬|자기가 한 일을 스스로 칭찬함
作心三日|작심삼일|결심이 사흘을 가지 못함
長幼有序|장유유서|어른과 아이 사이에는 차례가 있음
張三李四|장삼이사|평범한 보통 사람들
賊反荷杖|적반하장|도둑이 도리어 매를 듦: 잘못한 사람이 오히려 큰소리침
赤手空拳|적수공권|맨손과 맨주먹: 아무것도 없음
適者生存|적자생존|환경에 적응한 것만 살아남음
轉禍爲福|전화위복|재앙이 바뀌어 복이 됨
前代未聞|전대미문|지난 시대에는 들어 본 적 없음
前虎後狼|전호후랑|앞문의 호랑이 뒷문의 이리: 재앙이 계속 닥침
切磋琢磨|절차탁마|학문과 덕행을 갈고닦음
漸入佳境|점입가경|갈수록 경치나 일이 흥미로워짐
頂門一鍼|정문일침|따끔한 충고
正正堂堂|정정당당|태도나 수단이 바르고 떳떳함
糟糠之妻|조강지처|가난할 때 고생을 함께한 아내
朝令暮改|조령모개|법령을 자주 고쳐 갈피를 잡을 수 없음
朝變夕改|조변석개|계획이나 결정을 자주 바꿈
朝三暮四|조삼모사|눈앞의 이익만 알고 결과는 모름, 잔꾀로 남을 속임
左顧右眄|좌고우면|이쪽저쪽을 돌아보며 망설임
坐不安席|좌불안석|마음이 불안하여 한곳에 앉아 있지 못함
左之右之|좌지우지|제 마음대로 다룸
走馬看山|주마간산|달리는 말 위에서 산을 봄: 대충대충 훑어봄
走馬加鞭|주마가편|달리는 말에 채찍질: 잘하는 사람을 더욱 격려함
晝耕夜讀|주경야독|낮에는 농사 짓고 밤에는 공부함
晝夜長川|주야장천|밤낮으로 쉬지 않고 계속함
竹馬故友|죽마고우|어릴 때부터 같이 놀며 자란 벗
衆口難防|중구난방|많은 사람의 말을 막기 어려움
重言復言|중언부언|이미 한 말을 되풀이함
指鹿爲馬|지록위마|윗사람을 속이고 권세를 마음대로 함
知彼知己|지피지기|적과 나를 앎
至誠感天|지성감천|지극한 정성은 하늘도 감동시킴
知行合一|지행합일|지식과 행동이 일치함
進退兩難|진퇴양난|나아가기도 물러서기도 어려움
進退維谷|진퇴유곡|궁지에 몰려 꼼짝할 수 없음
此日彼日|차일피일|오늘내일 하며 자꾸 기한을 미룸
滄海一粟|창해일속|넓은 바다의 좁쌀 한 알: 아주 보잘것없는 존재
千變萬化|천변만화|끝없이 변화함
千辛萬苦|천신만고|온갖 어려움과 고생
天高馬肥|천고마비|하늘은 높고 말은 살찜: 가을을 이르는 말
天衣無縫|천의무봉|꾸밈없이 자연스러워 흠이 없음
千篇一律|천편일률|여럿이 모두 비슷비슷함
千載一遇|천재일우|천 년에 한 번 만날 만한 좋은 기회
千差萬別|천차만별|여러 가지 사물이 서로 다르고 차이가 많음
千萬多幸|천만다행|매우 다행함
天佑神助|천우신조|하늘과 신령의 도움
天地神明|천지신명|하늘과 땅의 모든 신령
天下泰平|천하태평|세상이 태평함: 아무 근심 없이 편안함
天眞爛漫|천진난만|꾸밈없이 순진하고 명랑함
靑山流水|청산유수|막힘없이 말을 잘함
靑天白日|청천백일|맑게 갠 대낮: 죄가 없음이 밝혀짐
靑出於藍|청출어람|제자가 스승보다 나음
草綠同色|초록동색|처지가 같은 것끼리 어울림
初志一貫|초지일관|처음 세운 뜻을 끝까지 밀고 나감
寸鐵殺人|촌철살인|짧은 말로 핵심을 찔러 감동시킴
忠言逆耳|충언역이|바른 말은 귀에 거슬림
醉生夢死|취생몽사|아무 뜻 없이 흐리멍덩하게 삶
針小棒大|침소봉대|작은 일을 크게 부풀려 말함
快刀亂麻|쾌도난마|어지러운 일을 명쾌하게 처리함
他山之石|타산지석|다른 사람의 하찮은 언행도 자기 수양에 도움이 됨
卓上空論|탁상공론|실현성 없는 헛된 논의
貪官汚吏|탐관오리|탐욕스럽고 부정한 관리
泰山北斗|태산북두|그 분야에서 가장 존경받는 사람
泰然自若|태연자약|어떤 일에도 태연하고 침착함
兎死狗烹|토사구팽|쓸모가 없어지면 버림받음
土崩瓦解|토붕와해|흙이 무너지고 기와가 깨짐: 조직이 걷잡을 수 없이 무너짐
破邪顯正|파사현정|그릇된 것을 깨고 바른 것을 드러냄
破竹之勢|파죽지세|대나무를 쪼개는 듯한 맹렬한 기세
八方美人|팔방미인|여러 방면에 능통한 사람, 누구에게나 호감을 사려는 사람
敗軍之將|패군지장|싸움에 진 장수
風前燈火|풍전등화|바람 앞의 등불: 매우 위태로운 처지
風餐露宿|풍찬노숙|바람 먹고 이슬에 잠: 고생스러운 객지 생활
匹夫之勇|필부지용|깊은 생각 없이 혈기로만 부리는 용기
必死卽生|필사즉생|죽기를 각오하고 싸우면 살아남
下石上臺|하석상대|아랫돌 빼서 윗돌 괴기: 임시변통
鶴首苦待|학수고대|목을 빼고 몹시 기다림
咸興差使|함흥차사|심부름 가서 돌아오지 않거나 소식이 없음
虛張聲勢|허장성세|실속 없이 큰소리치며 허세만 부림
懸頭刺股|현두자고|머리를 매달고 허벅지를 찌름: 열심히 공부함
螢雪之功|형설지공|고생하며 공부한 보람
狐假虎威|호가호위|남의 권세를 빌려 위세를 부림
好事多魔|호사다마|좋은 일에는 방해되는 일이 많음
虎視眈眈|호시탐탐|노려보며 기회를 엿봄
浩然之氣|호연지기|하늘과 땅 사이에 가득 찬 넓고 큰 기운
昏定晨省|혼정신성|저녁에는 잠자리를 보살피고 아침에는 문안을 드림
弘益人間|홍익인간|널리 인간 세상을 이롭게 함
畵龍點睛|화룡점정|가장 중요한 부분을 완성함
畵蛇添足|화사첨족|쓸데없는 것을 덧붙여 도리어 일을 그르침
花容月態|화용월태|꽃같이 고운 얼굴과 달같이 아름다운 몸매
花朝月夕|화조월석|꽃 피는 아침과 달 밝은 저녁: 경치 좋은 때
換骨奪胎|환골탈태|용모가 몰라보게 아름다워짐, 남의 글을 고쳐 새롭게 만듦
會者定離|회자정리|만난 사람은 반드시 헤어짐
興盡悲來|흥진비래|즐거운 일이 다하면 슬픈 일이 옴
興亡盛衰|흥망성쇠|흥하고 망하고 성하고 쇠함
喜怒哀樂|희로애락|기쁨과 노여움과 슬픔과 즐거움
兵家常事|병가상사|싸움에 이기고 지는 것은 흔한 일: 실패는 흔히 있는 일
本末顚倒|본말전도|일의 중요한 것과 덜 중요한 것이 뒤바뀜
先後本末|선후본말|먼저와 나중, 근본과 끝
首尾相應|수미상응|처음과 끝이 서로 맞음
面從腹背|면종복배|겉으로는 복종하는 체하고 속으로는 배반함
白衣民族|백의민족|흰옷을 입는 민족: 한민족을 이르는 말
父子有親|부자유친|아버지와 아들 사이에는 친애가 있어야 함
夫婦有別|부부유별|부부 사이에는 구별이 있음
朋友有信|붕우유신|친구 사이에는 믿음이 있어야 함
見利思義|견리사의|이익을 보면 의로운지 먼저 생각함
見利忘義|견리망의|이익을 보면 의리를 저버림
苛斂誅求|가렴주구|세금을 가혹하게 거두고 재물을 빼앗음
甘井先竭|감정선갈|물맛 좋은 우물이 먼저 마름: 재주 있는 사람이 먼저 쇠함
以身作則|이신작칙|자신이 먼저 본보기가 됨
小貪大失|소탐대실|작은 것을 탐내다 큰 것을 잃음
事實無根|사실무근|사실에 근거가 없음
敬老孝親|경로효친|노인을 공경하고 부모에게 효도함
兄友弟恭|형우제공|형은 아우를 사랑하고 아우는 형을 공경함
三從之道|삼종지도|여자가 따라야 할 세 가지 도리
天地開闢|천지개벽|하늘과 땅이 처음 열림: 세상이 크게 변함
天人共怒|천인공노|하늘과 사람이 함께 노할 만큼 나쁨
天生緣分|천생연분|하늘이 맺어 준 인연
水火相剋|수화상극|물과 불처럼 서로 맞서 해침
海枯石爛|해고석란|바다가 마르고 돌이 썩음: 아주 오랜 세월이 지나도 변치 않음
草木皆兵|초목개병|풀과 나무도 모두 적병으로 보임: 몹시 놀라 의심함
堂狗風月|당구풍월|서당 개 삼 년이면 풍월을 읊음: 오래 보고 들으면 익힘
後生可畏|후생가외|젊은 후배는 두려워할 만함
前途有望|전도유망|앞으로 잘될 가망이 있음
前人未踏|전인미답|아직 아무도 발을 들여놓지 못함
無事安逸|무사안일|큰 탈 없이 편안히 지냄: 일을 소극적으로 처리함
無事泰平|무사태평|아무 탈 없이 마음이 편안함
無知莫知|무지막지|몹시 무식하고 상스러우며 난폭함
無念無想|무념무상|아무 생각이 없는 상태
有始有終|유시유종|처음부터 끝까지 한결같음
自問自答|자문자답|스스로 묻고 스스로 대답함
自手成家|자수성가|물려받은 재산 없이 스스로 집안을 일으킴
自中之亂|자중지란|같은 편끼리의 싸움
自暴自棄|자포자기|스스로 자신을 포기하고 돌보지 않음
自由自在|자유자재|거침없이 마음대로 함
自力更生|자력갱생|자기 힘으로 일을 해 나감
自古以來|자고이래|옛날부터 지금까지
長生不死|장생불사|오래 살고 죽지 않음
利害得失|이해득실|이익과 손해, 얻음과 잃음
先憂後樂|선우후락|근심은 남보다 먼저, 즐거움은 남보다 나중에
生年月日|생년월일|태어난 해와 달과 날
生者必滅|생자필멸|살아 있는 것은 반드시 죽음
不老長生|불로장생|늙지 않고 오래 삶
不老不死|불로불사|늙지도 죽지도 않음
不可思議|불가사의|이해할 수 없는 신기한 일
不可抗力|불가항력|힘으로 어찌할 수 없는 일
不知其數|부지기수|그 수를 알 수 없을 만큼 많음
不遠千里|불원천리|천 리 길도 멀다 하지 않음
不立文字|불립문자|말이나 글로는 전할 수 없음
不眠不休|불면불휴|자지도 쉬지도 않음
氷肌玉骨|빙기옥골|아름다운 살결과 뼈대: 매화 또는 미인
事事件件|사사건건|모든 일마다
死生決斷|사생결단|죽고 사는 것을 가리지 않고 끝장을 냄
死而後已|사이후이|죽은 뒤에야 그만둠: 평생 힘씀
四海兄弟|사해형제|온 세상 사람이 모두 형제
四苦八苦|사고팔고|온갖 괴로움
三寒四溫|삼한사온|사흘 춥고 나흘 따뜻함
三日天下|삼일천하|권세를 잡았다가 곧 쫓겨남
上行下效|상행하효|윗사람이 하면 아랫사람이 본받음
相扶相助|상부상조|서로서로 도움
傷弓之鳥|상궁지조|화살에 다친 새: 한 번 혼난 뒤 몹시 겁을 냄
率先垂範|솔선수범|앞장서서 모범을 보임
手舞足蹈|수무족도|매우 기뻐서 날뜀
首鼠兩端|수서양단|거취를 정하지 못하고 망설임
隨機應變|수기응변|그때그때 형편에 맞게 처리함
十年知己|십년지기|오래전부터 친하게 사귄 친구
十常八九|십상팔구|열에 여덟아홉: 거의 틀림없음
十日之菊|십일지국|때가 지나 쓸모없게 된 것
深山幽谷|심산유곡|깊은 산속의 으슥한 골짜기
深謀遠慮|심모원려|깊은 꾀와 먼 앞날을 내다보는 생각
心機一轉|심기일전|어떤 계기로 마음을 완전히 바꿈
心心相印|심심상인|말없이 마음과 마음으로 뜻이 통함
阿鼻叫喚|아비규환|여러 사람이 비참한 지경에 빠져 울부짖는 상황
惡戰苦鬪|악전고투|매우 어렵고 괴로운 싸움
惡因惡果|악인악과|나쁜 일을 하면 나쁜 결과를 얻음
安分知足|안분지족|자기 분수를 지키며 만족할 줄 앎
眼高手低|안고수저|눈은 높으나 솜씨는 서투름
藥房甘草|약방감초|어떤 일에나 빠짐없이 끼어듦
養虎遺患|양호유환|호랑이를 길러 근심을 남김: 화근을 키움
魚魯不辨|어로불변|어(魚) 자와 노(魯) 자를 구별하지 못함: 매우 무식함
魚水之親|어수지친|물고기와 물처럼 친밀한 사이
與民同樂|여민동락|백성과 함께 즐김
如出一口|여출일구|한 입에서 나오는 것 같음: 여러 사람의 말이 같음
烏鳥私情|오조사정|부모를 봉양하는 효성
玉石俱焚|옥석구분|옥과 돌이 함께 탐: 좋은 것과 나쁜 것이 함께 망함
外親內疎|외친내소|겉으로는 친한 척하나 속으로는 멀리함
龍虎相搏|용호상박|강자끼리 서로 싸움
愚公移山|우공이산|우직하게 한 가지 일을 계속하면 결국 이룸
雨順風調|우순풍조|비와 바람이 알맞게 와서 농사에 좋음
遠水近火|원수근화|먼 데 있는 물은 가까운 불을 끄지 못함
流芳百世|유방백세|꽃다운 이름이 후세에 길이 전함
流言蜚語|유언비어|근거 없이 떠도는 소문
柔能制剛|유능제강|부드러운 것이 능히 강한 것을 이김
以德報怨|이덕보원|원수를 덕으로 갚음
以夷制夷|이이제이|적을 이용해 적을 막음
一葉知秋|일엽지추|작은 징조로 큰 흐름을 앎
一葉片舟|일엽편주|조그마한 조각배
一衣帶水|일의대수|한 줄기 띠 같은 좁은 강: 매우 가까운 사이
一絲不亂|일사불란|질서가 정연하여 조금도 어지러움이 없음
一瀉千里|일사천리|일이 거침없이 빨리 진행됨
一脈相通|일맥상통|생각이나 성질이 서로 통함
一刀兩斷|일도양단|어떤 일을 머뭇거리지 않고 단호히 결정함
一騎當千|일기당천|한 사람이 천 명을 당해 낼 만큼 힘이 셈
一朝一夕|일조일석|짧은 시일
一知半解|일지반해|하나쯤 알고 반쯤 이해함: 어설픈 지식
一筆揮之|일필휘지|단숨에 글씨나 그림을 써 내림
日進月步|일진월보|날로 달로 계속 발전함
日新月異|일신월이|나날이 새로워지고 다달이 달라짐
臨機應變|임기응변|그때그때 일어나는 사정에 맞게 처리함
臨戰無退|임전무퇴|싸움에 임하여 물러서지 않음
入鄕循俗|입향순속|그 지방에 가면 그곳 풍속을 따름
立身揚名|입신양명|출세하여 이름을 세상에 떨침
立身出世|입신출세|사회적으로 높은 자리에 오름
才色兼備|재색겸비|재주와 미모를 함께 갖춤
才勝薄德|재승박덕|재주는 뛰어나나 덕이 부족함
適材適所|적재적소|알맞은 인재를 알맞은 자리에 씀
電光石火|전광석화|매우 짧은 시간이나 재빠른 동작
戰戰兢兢|전전긍긍|몹시 두려워 벌벌 떨며 조심함
輾轉反側|전전반측|몸을 이리저리 뒤척이며 잠을 이루지 못함
全知全能|전지전능|모든 것을 알고 무엇이든 할 수 있음
切齒腐心|절치부심|몹시 분하여 이를 갈며 벼름
左衝右突|좌충우돌|이리저리 닥치는 대로 부딪침
主客顚倒|주객전도|주인과 손님의 위치가 뒤바뀜
酒池肉林|주지육림|호화스러운 술잔치
竹杖芒鞋|죽장망혜|간편한 나그네 차림
衆寡不敵|중과부적|적은 수로는 많은 수를 대적하지 못함
支離滅裂|지리멸렬|이리저리 흩어져 갈피를 잡을 수 없음
盡忠報國|진충보국|충성을 다하여 나라의 은혜에 보답함
天災地變|천재지변|자연 현상으로 일어나는 재앙
徹頭徹尾|철두철미|처음부터 끝까지 빈틈이 없음
靑雲之志|청운지지|높은 벼슬에 오르려는 큰 뜻
淸廉潔白|청렴결백|마음이 맑고 깨끗하며 재물 욕심이 없음
秋風落葉|추풍낙엽|세력이 갑자기 쇠퇴함
出將入相|출장입상|문무를 겸비한 사람
七顚八起|칠전팔기|실패해도 굴하지 않고 다시 일어남
太平聖代|태평성대|어질고 착한 임금이 다스리는 태평한 세상
破顔大笑|파안대소|얼굴이 활짝 펴지도록 크게 웃음
八面六臂|팔면육비|여러 방면에서 능숙하게 일을 처리함
平地風波|평지풍파|평온한 곳에서 일어나는 분쟁
抱腹絶倒|포복절도|배를 안고 넘어질 만큼 몹시 웃음
暴虎馮河|포호빙하|무모한 용기를 부림
風樹之嘆|풍수지탄|부모가 돌아가셔서 효도할 수 없는 슬픔
皮骨相接|피골상접|살가죽과 뼈가 맞닿을 정도로 몹시 야윔
下學上達|하학상달|낮은 데서 배워 높은 이치에 통달함
行雲流水|행운유수|떠가는 구름과 흐르는 물처럼 자연스러움
虛心坦懷|허심탄회|마음속에 거리낌 없이 솔직함
虛無孟浪|허무맹랑|터무니없이 거짓되고 실속이 없음
賢母良妻|현모양처|어진 어머니이자 착한 아내
懸河雄辯|현하웅변|흐르는 물처럼 거침없이 잘하는 말솜씨
好衣好食|호의호식|좋은 옷을 입고 좋은 음식을 먹음: 잘 삶
惑世誣民|혹세무민|세상 사람을 속여 어지럽힘
禍不單行|화불단행|불행은 연달아 옴
荒唐無稽|황당무계|말이나 행동이 터무니없고 근거가 없음
後悔莫及|후회막급|아무리 후회해도 소용없음
厚顔無恥|후안무치|얼굴이 두꺼워 부끄러움을 모름
興味津津|흥미진진|흥미가 넘쳐 흐름
喜喜樂樂|희희낙락|매우 기뻐하고 즐거워함
喜色滿面|희색만면|기쁜 빛이 얼굴에 가득함
角者無齒|각자무치|한 사람이 모든 복을 다 가질 수는 없음
"""


@st.cache_resource(show_spinner=False)
def idiom_db():
    """내장 사전 -> [{hanja, sound, meaning, group}] (4글자 사자성어)"""
    items, _problems, _dup = build_idioms(parse_idiom_text(IDIOM_DB_TEXT))
    return items


def word_chars():
    """단어장에 들어 있는 모든 한자(글자) 집합"""
    return {c for w in st.session_state.words for c in w["hanja"] if has_hanja(c)}


def match_idioms(chars):
    """내 한자로만 이루어진 사자성어"""
    return [dict(x) for x in idiom_db() if all(c in chars for c in x["hanja"])]


def near_idioms(chars):
    """딱 한 글자만 모르는 사자성어 -> [(사자성어, 모르는 한자)]"""
    out = []
    for x in idiom_db():
        miss = {c for c in x["hanja"] if c not in chars}
        if len(miss) == 1:
            out.append((x, next(iter(miss))))
    return out


def refresh_idioms():
    """단어장이 바뀌었을 때만 사자성어 목록을 다시 골라요 (내 한자 → 사자성어 자동 연결)"""
    s = st.session_state
    sig = "".join(sorted(word_chars()))
    if s.get("idiom_sig") != sig or not isinstance(s.idioms, list):
        s.idioms = match_idioms(set(sig))
        s.idiom_sig = sig
        reset_idiom_quiz()  # 목록이 바뀌었으니 진행 중이던 사자성어 시험은 비워요


# ----------------------------------------------------------------------
# 사자성어 (목록 · 시험 · 오답 노트) — 한자 통계와 섞지 않고 따로 저장해요
# ----------------------------------------------------------------------
def check_idiom_reading(user, it):
    """독음 채점: 띄어쓰기·'/' 등은 무시하고, '/'로 이은 여러 답 중 하나면 정답"""
    u = norm(user)
    return bool(u) and (u in split_alts(it["sound"]) or u == norm(it["sound"]))


def blank_text(h, bi):
    return h[:bi] + "□" + h[bi + 1:]


def masked_sound(it, bi):
    """독음에서도 같은 자리를 가려요 (글자 수가 맞을 때만)"""
    snd = re.sub(r"\s+", "", it["sound"])
    if len(snd) != len(it["hanja"]):
        return ""
    return snd[:bi] + "□" + snd[bi + 1:]


def idiom_char_pool():
    """빈칸 문제의 오답 글자 후보: 단어장 한자 + 사자성어 속 한자"""
    s = st.session_state
    pool = set()
    for w in s.words:
        pool.update(c for c in w["hanja"] if has_hanja(c))
    for x in s.idioms:
        pool.update(x["hanja"])
    return pool


def idiom_choices(it, field, n=4, avoid_sound=False):
    """보기 만들기. 뜻은 사전의 다른 사자성어 뜻을, 사자성어(한자)는 내 목록의 다른 사자성어를 먼저 쓰고,
    모자라면 한 글자를 바꾼 '가짜' 사자성어로 채워요 (실제 사자성어와 겹치지 않게)"""
    s = st.session_state
    me = it["hanja"]
    if field == "meaning":
        mine = [x for x in s.idioms if x["hanja"] != me]
        rest = [x for x in idiom_db() if x["hanja"] != me]
        random.shuffle(mine)
        random.shuffle(rest)
        seen, pick = {norm(it["meaning"])}, []
        for x in mine + rest:
            k = norm(x["meaning"])
            if k in seen:
                continue
            seen.add(k)
            pick.append(x)
            if len(pick) == n - 1:
                break
        opts = [{"key": x["hanja"], "text": x["meaning"]} for x in pick] + [{"key": me, "text": it["meaning"]}]
        random.shuffle(opts)
        return opts
    others = [x for x in s.idioms if x["hanja"] != me]
    if avoid_sound:  # 독음이 같은 다른 사자성어는 정답 시비가 생겨서 제외
        others = [x for x in others if norm(x["sound"]) != norm(it["sound"])]
    random.shuffle(others)
    texts, seen = [], {me}
    for x in others:
        if x["hanja"] not in seen:
            seen.add(x["hanja"])
            texts.append(x["hanja"])
            if len(texts) == n - 1:
                break
    chars = sorted(idiom_char_pool())
    real = {x["hanja"] for x in idiom_db()}
    tries = 0
    while len(texts) < n - 1 and tries < 300 and len(chars) >= 2:
        tries += 1
        pos, c = random.randrange(len(me)), random.choice(chars)
        fake = me[:pos] + c + me[pos + 1:]
        if c == me[pos] or fake in seen or fake in real:
            continue
        seen.add(fake)
        texts.append(fake)
    opts = [{"key": t, "text": t} for t in texts] + [{"key": me, "text": me}]
    random.shuffle(opts)
    return opts


def blank_choices(it, bi, cpool=None, n=4):
    ans = it["hanja"][bi]
    pool = sorted((cpool if cpool is not None else idiom_char_pool()) - set(it["hanja"]))
    random.shuffle(pool)
    opts = pool[: n - 1] + [ans]
    random.shuffle(opts)
    return [{"key": c, "text": c} for c in opts]


def make_idiom_item(it, ty, cpool=None):
    it = dict(it)
    q = {"it": it, "type": ty}
    if ty == "i_ch_snd":
        q["opts"] = idiom_choices(it, "hanja", avoid_sound=True)
    elif ty == "i_ch_hj":
        q["opts"] = idiom_choices(it, "hanja")
    elif ty == "i_ch_ko":
        q["opts"] = idiom_choices(it, "meaning")
    elif ty in ("i_blank", "i_hw"):
        q["bi"] = random.randrange(len(it["hanja"]))
        if ty == "i_blank":
            q["opts"] = blank_choices(it, q["bi"], cpool)
    if "opts" in q and len(q["opts"]) < 2:  # 보기를 못 만들면 독음 쓰기로 대신
        q = {"it": it, "type": "i_read"}
    return q


def idiom_kinds_avail():
    """(고르기 가능, 빈칸 고르기 가능, 빈칸 손글씨 가능)"""
    n, pool = len(st.session_state.idioms), len(idiom_char_pool())
    choice_ok = n >= 1 and pool >= 5
    blank_ok = choice_ok and pool >= 10
    hw_ok = bool(HAS_CANVAS and find_font())
    return choice_ok, blank_ok, hw_ok


def idiom_types_for(kind, avail=None):
    choice_ok, blank_ok, hw_ok = avail or idiom_kinds_avail()
    ok = {"i_read": True, "i_ch_snd": choice_ok, "i_ch_ko": choice_ok, "i_ch_hj": choice_ok,
          "i_blank": blank_ok, "i_hw": hw_ok}
    if kind == "섞어서":
        return [t for t in IDIOM_LABELS if ok[t]]
    return [t for t in IDIOM_KIND_TYPES.get(kind, []) if ok[t]]


def idiom_q_text(q):
    it, ty = q["it"], q["type"]
    if ty in ("i_read", "i_ch_ko"):
        return it["hanja"]
    if ty == "i_ch_snd":
        return it["sound"]
    if ty == "i_ch_hj":
        return it["meaning"]
    return blank_text(it["hanja"], q["bi"])


def idiom_a_text(q):
    it, ty = q["it"], q["type"]
    table = {"i_read": it["sound"], "i_ch_snd": it["hanja"], "i_ch_ko": it["meaning"], "i_ch_hj": it["hanja"]}
    return table.get(ty) or it["hanja"][q["bi"]]


def record_idiom(q, ok, ans, retry=False):
    """사자성어 풀이 기록 (한자 통계와 따로 idiom_stats / idiom_wrong 에 저장)"""
    s = st.session_state
    it = q["it"]
    key = it["hanja"]
    e = s.idiom_stats.setdefault(key, {"sound": it["sound"], "meaning": it["meaning"], "right": 0, "wrong": 0})
    if ok:
        e["right"] += 1
        if retry and key in s.idiom_wrong:
            del s.idiom_wrong[key]
    else:
        e["wrong"] += 1
        w = s.idiom_wrong.setdefault(
            key, {"hanja": key, "sound": it["sound"], "meaning": it["meaning"], "wrong_count": 0, "last": ""})
        w["wrong_count"] += 1
        w["last"] = datetime.now(KST).strftime("%Y-%m-%d %H:%M")
    return {"q": q, "correct": bool(ok), "answer": str(ans)}


# --- 손글씨 공통 (사자성어 빈칸용; 연습 탭의 render_write는 그대로 둬요) -------
def canvas_backdrop_for(box_key, img, px):
    """canvas_backdrop 과 같지만 칸 이름(box_key)을 따로 받아서, 여러 화면의 캔버스가 서로 겹치지 않게 해요"""
    buf = io.BytesIO()
    img.save(buf, "PNG")
    uri = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
    html(
        f"<style>.st-key-{box_key}{{width:{px + 52}px; max-width:100%; background:url(" + uri + ") 0 0 / "
        f"{px}px {px}px no-repeat;}}"
        f".st-key-{box_key} iframe{{mix-blend-mode:multiply;}}</style>"
    )


def hw_candidates():
    s = st.session_state
    chars = [w["hanja"] for w in s.words if len(w["hanja"]) == 1]
    for x in s.idioms:
        chars.extend(x["hanja"])
    return list(dict.fromkeys(chars))[:800]


def hw_grade(res, ch, font, cands):
    """캔버스 결과를 기존 grade_writing 으로 조용히 채점 -> (맞았나, 표시용 답)"""
    data = None
    if res is not None and getattr(res, "image_data", None) is not None:
        data = flatten(res.image_data)
    if data is None or (data[:, :, :3].astype(float).mean(axis=2) < 110).sum() < 30:
        return False, "(빈칸)"
    if not font:
        return False, "(채점 불가)"
    r = grade_writing(data, ch, font, grade_params(), cands)
    if r is None:
        return False, "(인식 못함)"
    return bool(r["ok"]), f"유사도 {r['score']:.2f}"


def hw_clear():
    st.session_state.hw_clear += 1


# --- 문제 그리기 (사자성어 시험) -------------------------------------
def idiom_card(text):
    n = max(1, len(text))
    size = 64 if n <= 4 else 50 if n <= 6 else 40
    html(f"<div class='hj-card wide' style='font-size:min({size}px,{int(80 / n)}vw)'>{esc(text)}</div>")


def _idiom_pick(q, o, answer_fn):
    it, ty = q["it"], q["type"]
    right = it["hanja"][q["bi"]] if ty == "i_blank" else it["hanja"]
    answer_fn(o["key"] == right, o["text"])


def _idiom_skip(answer_fn):
    answer_fn(False, "(모르겠어요)")


def render_idiom_q(q, scope, answer_fn, pos, nonce):
    """사자성어 문제 한 개. 맞고 틀림은 answer_fn(ok, 내 답)으로 넘기고, 정답은 보여주지 않아요"""
    s = st.session_state
    it, ty = q["it"], q["type"]
    base = f"{scope}_{pos}_{nonce}"
    if ty == "i_read":
        idiom_card(it["hanja"])
        html("<div class='ask'>이 사자성어의 <b>독음</b>을 입력하세요. 예) 일석이조</div>")
        with st.form(f"if_{base}", border=False):
            ans = st.text_input("독음", key=f"ia_{base}", label_visibility="collapsed", placeholder="독음")
            c1, c2 = st.columns(2)
            go = c1.form_submit_button("다음 ➡️", type="primary", width="stretch")
            skip = c2.form_submit_button("모르겠어요", width="stretch")
        if go or skip:
            a = (ans or "").strip() if go else ""
            answer_fn(check_idiom_reading(a, it), a or ("(빈칸)" if go else "(모르겠어요)"))
            st.rerun()
        return
    if ty in ("i_blank", "i_hw"):
        bi = q["bi"]
        ms = masked_sound(it, bi)
        clue = f"뜻: {esc(it['meaning'])}" + (f" · 독음: {esc(ms)}" if ms else "")
    if ty == "i_hw":
        px, font = canvas_px(), find_font()
        left, right = st.columns([3, 2])
        with left:
            idiom_card(blank_text(it["hanja"], bi))
            html(f"<div class='ask'>{clue}<br>□에 들어갈 한자를 오른쪽 칸에 쓰세요</div>")
        canvas_backdrop_for(f"canvas_box_{scope}", grid_image(px), px)
        with right:
            with st.container(key=f"canvas_box_{scope}"):
                res = draw_canvas(f"icv_{base}_{s.hw_clear}_{px}", px)
        b1, b2, b3 = st.columns(3)
        if b1.button("다음 ➡️", key=f"ih_go_{base}", type="primary", width="stretch"):
            ok, shown = hw_grade(res, it["hanja"][bi], font, hw_candidates())
            answer_fn(ok, shown)
            st.rerun()
        b2.button("지우기", key=f"ih_clr_{base}", on_click=hw_clear, width="stretch")
        b3.button("모르겠어요", key=f"ih_skip_{base}", on_click=_idiom_skip, args=(answer_fn,), width="stretch")
        return
    # 고르기 유형
    if ty == "i_ch_snd":
        html(f"<div class='mean-card'><b>{esc(it['sound'])}</b></div>")
        html("<div class='ask'>이 독음에 해당하는 <b>사자성어</b>는?</div>")
        cls = "hj"
    elif ty == "i_ch_ko":
        idiom_card(it["hanja"])
        html("<div class='ask'>이 사자성어의 <b>뜻</b>은?</div>")
        cls = "text"
    elif ty == "i_ch_hj":
        html(f"<div class='mean-card' style='font-size:1.25rem'>{esc(it['meaning'])}</div>")
        html("<div class='ask'>이 뜻에 해당하는 <b>사자성어</b>는?</div>")
        cls = "hj"
    else:  # i_blank
        idiom_card(blank_text(it["hanja"], bi))
        html(f"<div class='ask'>{clue}<br>□에 들어갈 <b>한자</b>는?</div>")
        cls = "blank"
    wide = cls == "text" or (cls == "hj" and max(len(o["text"]) for o in q["opts"]) > 5)
    with st.container(key=f"xc_{cls}_{scope}"):
        cols = st.columns(1 if wide else 2)
        for i, o in enumerate(q["opts"]):
            cols[i % len(cols)].button(o["text"], key=f"io_{base}_{i}", on_click=_idiom_pick,
                                       args=(q, o, answer_fn), width="stretch")
    st.button("모르겠어요", key=f"is_{base}", on_click=_idiom_skip, args=(answer_fn,))


# --- 사자성어 시험 ---------------------------------------------------------------------
def idiom_pool_for(src):
    s = st.session_state
    if src == "오답 노트만":
        return [{"hanja": v["hanja"], "sound": v["sound"], "meaning": v["meaning"], "group": ""}
                for v in s.idiom_wrong.values() if v.get("sound") and v.get("meaning")]
    return [dict(x) for x in s.idioms]


def start_idiom_quiz():
    s = st.session_state
    src, kind = s.get("idm_src", "전체"), s.get("idm_type", "섞어서")
    pool = idiom_pool_for(src)
    if not pool:
        s.flash = ("오답 노트에 사자성어가 없어요. 🎉" if src == "오답 노트만"
                   else "내 한자로 만들 수 있는 사자성어가 아직 없어요.")
        return
    types = idiom_types_for(kind)
    if not types:
        s.flash = "고른 유형으로는 문제를 만들 수 없어요. (고르기는 사자성어가 4개 이상, 손글씨는 캔버스·폰트가 필요해요)"
        return
    random.shuffle(pool)
    cpool = idiom_char_pool()
    count = _int(s.get("idm_count"), 10, 1, 100)
    items = [make_idiom_item(it, random.choice(types), cpool) for it in pool[:count]]
    reset_idiom_quiz()
    s.iquiz, s.iretry, s.flash = items, (src == "오답 노트만"), ""


def idiom_answer(ok, shown):
    s = st.session_state
    if s.ipos >= len(s.iquiz):
        return
    q = s.iquiz[s.ipos]
    s.ires.append(record_idiom(q, ok, shown, retry=s.iretry))
    s.ipos += 1
    s.inonce += 1


def idiom_retry_wrong():
    s = st.session_state
    items = [make_idiom_item(r["q"]["it"], r["q"]["type"]) for r in s.ires if not r["correct"]]
    if items:
        reset_idiom_quiz()
        s.iquiz, s.iretry = items, False


def render_idiom_setup():
    s = st.session_state
    n = len(s.idioms)
    st.subheader("사자성어 시험")
    st.caption(f"내 한자로 만든 사자성어 {n}개에서 나와요 · 한자 시험과 따로 기록해요. 정답은 끝난 뒤에 보여줘요.")
    av = idiom_kinds_avail()
    kinds = [k for k in IDIOM_KINDS if idiom_types_for(k, av)]
    if n and av[0] and "빈칸 손글씨" not in kinds:
        st.caption("'빈칸 손글씨'는 손글씨 기능과 한자 폰트가 있어야 해요. (지금은 빠져 있어요)")
    if s.get("idm_type") not in kinds:
        s.idm_type = kinds[0]
    st.radio("유형", kinds, key="idm_type", horizontal=True)
    if s.get("idm_src") not in IDIOM_SRC:
        s.idm_src = IDIOM_SRC[0]
    st.radio("범위", IDIOM_SRC, key="idm_src", horizontal=True)
    st.slider("문제 수", 5, 50, step=5, key="idm_count")
    st.button("▶️ 시작하기", type="primary", on_click=start_idiom_quiz, width="stretch", key="idm_start")


def render_idiom_result():
    s = st.session_state
    res = s.ires
    n = len(res)
    c = sum(r["correct"] for r in res)
    st.subheader("시험 결과예요")
    rate = c / n * 100 if n else 0
    html(f"<div class='score-box'><div class='big'>{c} / {n}</div>"
         f"<div class='lbl'>정답률 {rate:.0f}% · 틀린 것은 사자성어 오답 노트에 담았어요</div></div>")
    wrong = [r for r in res if not r["correct"]]
    if wrong:
        st.markdown("##### 틀린 문제")
        st.dataframe(pd.DataFrame([{
            "유형": IDIOM_LABELS[r["q"]["type"]], "문제": idiom_q_text(r["q"]), "내 답": r["answer"],
            "정답": idiom_a_text(r["q"]), "사자성어": r["q"]["it"]["hanja"], "독음": r["q"]["it"]["sound"],
            "뜻": r["q"]["it"]["meaning"]} for r in wrong]), width="stretch", hide_index=True)
    elif n:
        st.success("전부 맞혔어요! 🎉")
    b1, b2 = st.columns(2)
    b1.button("🔁 새로 시작", on_click=start_idiom_quiz, width="stretch", key="idm_again")
    if wrong:
        b2.button("❌ 틀린 것만 다시", on_click=idiom_retry_wrong, width="stretch", key="idm_wrong_again")
    st.button("설정 바꾸기", on_click=reset_idiom_quiz, key="idm_back")


def render_idiom_exam():
    s = st.session_state
    if not s.idioms:
        st.info("지금 단어장의 한자만으로 만들 수 있는 사자성어가 아직 없어요. 한자를 더 불러오면 자동으로 늘어나요.")
        st.button("📜 목록 보기", on_click=go_idiom_list, key="idm_go_list")
        return
    if not s.iquiz:
        render_idiom_setup()
        return
    if s.ipos >= len(s.iquiz):
        render_idiom_result()
        return
    q = s.iquiz[s.ipos]
    st.progress(s.ipos / len(s.iquiz), text=f"{s.ipos + 1} / {len(s.iquiz)}  ·  {IDIOM_LABELS[q['type']]}")
    render_idiom_q(q, "idm", idiom_answer, s.ipos, s.inonce)
    st.button("그만하기", key="idm_quit", on_click=reset_idiom_quiz)


# --- 사자성어 오답 노트 -----------------------------------------------------------------
def idiom_wrong_delete():
    s = st.session_state
    for k in s.get("idm_wrong_sel", []):
        s.idiom_wrong.pop(k, None)
    s.idm_wrong_sel = []


def idiom_wrong_clear():
    st.session_state.idiom_wrong.clear()


def go_idiom_wrong_quiz():
    s = st.session_state
    s.idm_src, s.idm_type, s.idm_view = "오답 노트만", "섞어서", "시험"
    start_idiom_quiz()


def render_idiom_wrong():
    s = st.session_state
    if not s.idiom_wrong:
        st.info("사자성어 오답 노트가 비어 있어요. 틀린 사자성어가 여기에 쌓여요.")
        return
    df = pd.DataFrame([{"사자성어": v["hanja"], "독음": v["sound"], "뜻": v["meaning"],
                        "틀린 횟수": v["wrong_count"], "최근": v["last"]} for v in s.idiom_wrong.values()]
                      ).sort_values("틀린 횟수", ascending=False)
    st.dataframe(df, width="stretch", hide_index=True)
    s.idm_wrong_sel = [x for x in s.get("idm_wrong_sel", []) if x in s.idiom_wrong]
    st.multiselect("오답 노트에서 뺄 사자성어", list(s.idiom_wrong.keys()), key="idm_wrong_sel")
    c1, c2, c3 = st.columns(3)
    c1.button("선택 삭제", on_click=idiom_wrong_delete, width="stretch", key="idm_wrong_del")
    c2.button("전체 비우기", on_click=idiom_wrong_clear, width="stretch", key="idm_wrong_all")
    c3.download_button("CSV로 저장", df.to_csv(index=False).encode("utf-8-sig"), "사자성어_오답노트.csv",
                       "text/csv", width="stretch", key="idm_wrong_csv")
    st.button("📝 오답으로 바로 시험 보기", type="primary", on_click=go_idiom_wrong_quiz, width="stretch",
              key="idm_wrong_quiz")
    st.caption("오답 노트 범위에서 맞히면 노트에서 자동으로 빠져요.")


# --- 사자성어 목록 (내 한자로 만든 것) ----------------------------------------------------
def go_idiom_list():
    st.session_state.idm_view = "목록"


def _idiom_row(x, miss=None):
    h = esc(x["hanja"])
    if miss:
        h = h.replace(esc(miss), f"<b class='miss'>{esc(miss)}</b>")
    return (f"<div class='idm'><span class='idm-h'>{h}</span> <span class='idm-s'>{esc(x['sound'])}</span>"
            f"<div class='idm-m'>{esc(x['meaning'])}</div></div>")


def render_idiom_list():
    s = st.session_state
    chars = word_chars()
    n = len(s.idioms)
    html(f"<div class='chips'><span class='chip good'>📜 내 한자로 만든 사자성어 {n}개</span>"
         f"<span class='chip'>🔤 내 한자 {len(chars)}개</span>"
         f"<span class='chip'>사전 {len(idiom_db())}개 중</span></div>")
    if not n:
        st.info("지금 단어장의 한자만으로 만들 수 있는 사자성어가 아직 없어요. 한자를 더 불러오면 자동으로 늘어나요.")
    else:
        c1, c2 = st.columns([3, 2])
        q = c1.text_input("찾기", key="idm_find", placeholder="사자성어·독음·뜻 검색", label_visibility="collapsed")
        hide = c2.toggle("독음·뜻 가리기", key="idm_hide", help="사자성어만 보이고, 눌러서 독음과 뜻을 확인해요 (암기용)")
        terms = [t for t in re.split(r"\s+", (q or "").strip()) if t]
        shown = [x for x in s.idioms
                 if all(t in x["hanja"] or t in x["sound"] or t in x["meaning"] for t in terms)]
        st.caption(f"{len(shown)}개 표시 중 · 이 목록의 사자성어가 '시험'에 나와요.")
        rows = []
        for x in shown:
            if hide:
                rows.append(f"<details class='idm'><summary><span class='idm-h'>{esc(x['hanja'])}</span></summary>"
                            f"<div class='idm-s'>{esc(x['sound'])}</div><div class='idm-m'>{esc(x['meaning'])}</div></details>")
            else:
                rows.append(_idiom_row(x))
        html("".join(rows))
    near = near_idioms(chars)
    if near:
        near.sort(key=lambda t: t[0]["sound"])
        with st.expander(f"🔎 한 글자만 더 알면 되는 사자성어 ({len(near)}개)"):
            st.caption("빨간 한자만 익히면 이 사자성어가 내 목록에 들어와요.")
            html("".join(_idiom_row(x, m) for x, m in near[:60]))
            if len(near) > 60:
                st.caption(f"… 외 {len(near) - 60}개")


def render_idioms():
    s = st.session_state
    views = ["목록", "시험", "오답 노트"]
    if s.get("idm_view") not in views:
        s.idm_view = views[0]
    st.radio("사자성어 화면", views, key="idm_view", horizontal=True, label_visibility="collapsed")
    if s.idm_view == "오답 노트":
        render_idiom_wrong()
    elif s.idm_view == "시험":
        render_idiom_exam()
    else:
        render_idiom_list()


# ----------------------------------------------------------------------
# 시험 날짜 기반 학습 플랜 (연습 탭 맨 위 카드) — 모두 '참고용' 계산이에요
# ----------------------------------------------------------------------
def exam_dday():
    """시험까지 남은 일수 (날짜가 없으면 None, 지났으면 음수)"""
    d = _iso(st.session_state.get("set_exam_date", ""))
    return (date.fromisoformat(d) - today_date()).days if d else None


def plan_words():
    s = st.session_state
    groups = s.get("plan_groups", [])
    return [w for w in s.words if not groups or w.get("group") in groups]


def recent_accuracy(words, days=14):
    """최근 정답률. 최근 14일 푼 문제가 5개 미만이면 범위 한자의 누적 정답률로 대신해요"""
    s = st.session_state
    r = w = 0
    for i in range(days):
        h = s.history.get((today_date() - timedelta(days=i)).isoformat())
        if h:
            r += h.get("right", 0)
            w += h.get("wrong", 0)
    if r + w >= 5:
        return r / (r + w), "최근 14일"
    r = sum(s.stats.get(x["hanja"], {}).get("right", 0) for x in words)
    w = sum(s.stats.get(x["hanja"], {}).get("wrong", 0) for x in words)
    return (r / (r + w) if r + w else 0.0), "누적"


def plan_numbers():
    """플랜 계산 결과 dict (시험일이 없거나 지났으면 None)"""
    s = st.session_state
    D = exam_dday()
    if D is None or D < 0:
        return None
    words = plan_words()
    unseen = sum(1 for w in words if w["hanja"] not in s.srs)
    new_today = 0 if D <= 2 else min(unseen, math.ceil(unseen / max(1, D - 2)))
    _n, review = due_counts(words)
    seen_ratio = (len(words) - unseen) / len(words) if words else 0.0
    acc, acc_src = recent_accuracy(words)
    return {"D": D, "total": len(words), "unseen": unseen, "new_today": new_today, "review": review,
            "seen_ratio": seen_ratio, "acc": acc, "acc_src": acc_src,
            "expected": round(100 * seen_ratio * acc)}


def start_plan():
    """오늘 복습 + 오늘 새 한자를 묶어서 기존 start_quiz 흐름으로 시작"""
    s = st.session_state
    p = plan_numbers()
    if not p:
        return
    pool = due_pool(plan_words())
    review = [w for w in pool if w["hanja"] in s.srs]
    new = [w for w in pool if w["hanja"] not in s.srs][: p["new_today"]]
    items = review + new
    if not items:
        s.flash = "오늘 플랜에 넣을 한자가 없어요. 내일 다시 만나요! 🎉"
        return
    start_quiz(pool_override=items, count_override=min(len(items), 200))


def clear_exam_date():
    st.session_state.set_exam_date = ""


def render_plan_card():
    s = st.session_state
    p = plan_numbers()
    D = exam_dday()
    with st.container(border=True):
        st.markdown("##### 📅 시험 플랜")
        if p is None:
            if D is not None:
                st.info("시험일이 지났어요. 새 시험 날짜를 정하면 플랜이 다시 만들어져요.")
            else:
                st.caption("시험 날짜와 목표 점수를 정하면, 오늘 새로 볼 한자와 복습할 한자를 계산해 드려요.")
        elif p["total"] == 0:
            st.info("플랜 범위에 한자가 없어요. 아래 설정에서 범위를 바꿔 주세요.")
        else:
            dd = "D-Day" if p["D"] == 0 else f"D-{p['D']}"
            target = _int(s.get("set_target"), 90, 0, 100)
            chips = [f"<span class='chip hot'>{dd}</span>",
                     f"<span class='chip'>🌱 오늘 새 한자 {p['new_today']}개</span>",
                     f"<span class='chip'>🔁 오늘 복습 {p['review']}개</span>",
                     f"<span class='chip'>안 본 한자 {p['unseen']}개</span>"]
            html("<div class='chips'>" + "".join(chips) + "</div>")
            html(f"<div class='score-box'><div class='big'>예상 {p['expected']}점 / 목표 {target}점</div>"
                 f"<div class='lbl'>참고용 · 본 한자 {p['seen_ratio']:.0%} × 최근 정답률 {p['acc']:.0%}</div></div>")
            if p["D"] <= 2:
                st.caption("시험 직전 이틀은 새 한자 없이 복습만 해요.")
            n_today = p["review"] + p["new_today"]
            st.button(f"▶️ 오늘 플랜 시작 (복습 {p['review']} + 새 한자 {p['new_today']})", type="primary",
                      on_click=start_plan, disabled=(n_today == 0), width="stretch", key="plan_start")
            if n_today == 0:
                st.caption("오늘 할 몫을 다 했어요. 🎉")
            st.caption("계산 방법(참고용): 오늘 새 한자 = 안 본 한자 ÷ (남은 날 − 2)를 올림 (마지막 이틀은 복습 기간) · "
                       "예상 점수 = 100 × 본 한자 비율 × 최근 정답률 · 오늘 복습 = 복습 날짜가 된 한자")
        with st.expander("⚙️ 플랜 설정", expanded=(D is None)):
            with st.form("plan_form", border=False):
                cur = _iso(s.get("set_exam_date", ""))
                today = today_date()
                value = max(date.fromisoformat(cur), today) if cur else today + timedelta(days=14)
                d = st.date_input("시험 날짜", value=value, min_value=today, format="YYYY/MM/DD")
                st.number_input("목표 점수 (0~100)", 0, 100, step=5, key="set_target")
                groups = sorted({w["group"] for w in s.words if w.get("group")},
                                key=lambda g: [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", g)])
                if groups:
                    s.plan_groups = [g for g in s.get("plan_groups", []) if g in groups]
                    st.multiselect("플랜 범위 분류 (비워 두면 전체)", groups, key="plan_groups",
                                   placeholder="전체 (눌러서 고르기)")
                saved = st.form_submit_button("💾 플랜 저장", type="primary", width="stretch")
            if saved:
                s.set_exam_date = d.isoformat()
                st.toast("플랜을 저장했어요.")
                st.rerun()
            if cur:
                st.button("시험 날짜 지우기", on_click=clear_exam_date, key="plan_clear")


# --- 단어장 탭: 별표·메모 / 내장 단어장 / 편집 ---------------------------------------
def render_note_tools(words):
    s = st.session_state
    n_star = sum(1 for e in s.notes.values() if e.get("star"))
    n_memo = sum(1 for e in s.notes.values() if e.get("note"))
    with st.expander("⭐ 별표·메모"):  # 제목을 고정해야 눌러도 접히지 않아요
        st.caption(f"지금 ★ {n_star}개 · 메모 {n_memo}개")
        if not words:
            st.caption("지금 보이는 한자가 없어요. 위 검색을 지우고 다시 해 보세요.")
            return
        by = {w["hanja"]: w for w in s.words}
        opts = [w["hanja"] for w in words][:3000]

        def fmt(h):
            w = by.get(h, {})
            return f"{'★ ' if is_star(h) else ''}{h}  {w.get('meaning', '')} {w.get('sound', '')}".strip()

        if s.get("nt_pick") not in opts:
            s.nt_pick = opts[0]
        h = st.selectbox("한자 고르기 (위 검색으로 좁힐 수 있어요)", opts, key="nt_pick", format_func=fmt)
        star = is_star(h)
        st.button("★ 별표 빼기" if star else "☆ 별표 주기", key="nt_star", on_click=toggle_star,
                  args=(h, not star), width="stretch")
        with st.form("nt_form", border=False):
            txt = st.text_area(f"'{h}' 내 메모 (최대 {MAX_NOTE}자)", value=note_of(h), max_chars=MAX_NOTE,
                               key=f"nt_txt_{h}", height=90)
            saved = st.form_submit_button("💾 메모 저장", type="primary", width="stretch")
        if saved:
            save_note(h, txt)
        s.nt_bulk = [x for x in s.get("nt_bulk", []) if x in opts]
        with st.form("nt_bulk_form", border=False):
            st.multiselect("별표를 한꺼번에 주거나 뺄 한자", opts, key="nt_bulk", format_func=fmt,
                           placeholder="여러 개 고르기")
            b1, b2 = st.columns(2)
            b1.form_submit_button("★ 별표 주기", on_click=bulk_star, args=(True,), width="stretch")
            b2.form_submit_button("☆ 별표 빼기", on_click=bulk_star, args=(False,), width="stretch")
        st.caption("별표를 준 한자는 연습·카드의 '별표 한자만' 범위로 풀 수 있어요.")


def render_word_edit():
    s = st.session_state
    with st.expander("✏️ 단어장 편집 (추가 · 수정 · 삭제)"):
        st.caption("저장하면 단어장이 바뀌고 이 기기에 자동 저장돼요. 진행 중이던 퀴즈·카드는 비워져요.")
        ver = s.ed_ver
        st.markdown("**한 개 추가**")
        with st.form("word_add_form", border=False):
            c1, c2 = st.columns(2)
            h = c1.text_input("한자", key=f"wa_h_{ver}")
            sd = c2.text_input("음", key=f"wa_s_{ver}")
            c3, c4 = st.columns(2)
            m = c3.text_input("뜻", key=f"wa_m_{ver}")
            g = c4.text_input("분류 (선택)", key=f"wa_g_{ver}")
            add = st.form_submit_button("➕ 한자 추가", type="primary", width="stretch")
        if add:
            h, m, sd, g = re.sub(r"\s+", "", h or ""), (m or "").strip(), (sd or "").strip(), (g or "").strip()
            why = word_problem(h, m, sd, g)
            if not why and any(w["hanja"] == h for w in s.words):
                why = "이미 있는 한자예요. 아래 표에서 고쳐 주세요."
            if why:
                st.warning(why)
            else:
                commit_words(s.words + [{"hanja": h, "meaning": m, "sound": sd, "group": g}])
                st.rerun()
        st.markdown("**표로 고치기**")
        st.caption("칸을 눌러 고치고, 표 맨 아래 빈 줄에서 새 줄을 더하고, 줄 왼쪽 칸을 눌러 고른 뒤 휴지통(또는 Delete 키)으로 지워요. "
                   "다 고쳤으면 '저장'을 눌러 주세요.")
        with st.form("word_edit_form", border=False):
            edited = st.data_editor(
                words_df(s.words), num_rows="dynamic", hide_index=True, key=f"word_editor_{ver}",
                column_config={
                    "한자": st.column_config.TextColumn("한자", max_chars=12),
                    "뜻": st.column_config.TextColumn("뜻", max_chars=100),
                    "음": st.column_config.TextColumn("음", max_chars=60),
                    "분류": st.column_config.TextColumn("분류", max_chars=30),
                })
            saved = st.form_submit_button("💾 단어장 저장", type="primary", width="stretch")
        if saved:
            new, errs = words_from_df(edited)
            if errs:
                st.error("저장하지 못했어요. 아래를 고친 뒤 다시 눌러 주세요.\n\n"
                         + "\n".join(f"- {e}" for e in errs[:8])
                         + (f"\n- … 외 {len(errs) - 8}건" if len(errs) > 8 else ""))
            elif not new:
                st.error("한자가 하나도 없으면 저장할 수 없어요.")
            elif new == s.words:
                st.info("바뀐 내용이 없어요.")
            else:
                commit_words(new)
                st.rerun()
        st.download_button("📥 현재 단어장 CSV 받기", words_csv(s.words), "내_단어장.csv", "text/csv", width="stretch")
        st.caption("받은 파일을 GitHub 저장소의 wordsets/ 폴더에 올리면 '내장 단어장'이 돼요. 파일 이름이 세트 이름이에요.")


def render_word_paste():
    s = st.session_state
    with st.expander("➕ 붙여넣어서 한자 추가"):
        st.caption("지금 단어장은 그대로 두고, 한자를 뒤에 더해요. 한 줄에 '한자 뜻 음' (예: 天 하늘 천). "
                   "쉼표·탭으로 나누면 4번째 칸은 분류로 들어가요. 이미 있는 한자는 새 뜻·음만 '/'로 덧붙여요.")
        st.text_area("추가할 한자", height=140, key="wp_text", placeholder="天 하늘 천\n地 땅 지")
        st.button("➕ 단어장에 추가", type="primary", on_click=paste_add_cb, args=("wp_text",),
                  key="wp_add_btn", width="stretch")
        st.caption("추가하면 진행 중이던 퀴즈·카드는 비워져요.")


def render_word_delete():
    s = st.session_state
    with st.expander("🗑️ 한자 삭제"):
        by = {w["hanja"]: w for w in s.words}
        opts = list(by)[:3000]
        s.del_pick = [x for x in s.get("del_pick", []) if x in by]  # 단어장이 바뀌어도 위젯이 깨지지 않게

        def fmt(h):
            w = by.get(h, {})
            return f"{h}  {w.get('meaning', '')} {w.get('sound', '')}".strip()

        st.multiselect("지울 한자 고르기", opts, key="del_pick", format_func=fmt, placeholder="여러 개 고르기")
        st.checkbox("그 한자의 오답·통계·복습·메모 기록도 함께 지우기", key="del_rec",
                    help="체크하지 않으면 단어장에서만 빠지고, 이미 쌓인 기록은 남아요.")
        st.button("선택한 한자 삭제", on_click=delete_selected_cb, key="del_sel_btn", width="stretch",
                  disabled=not s.get("del_pick"))
        st.divider()
        st.markdown("**단어장 전체 삭제**")
        st.caption(f"한자 {len(s.words)}개를 모두 지우고 처음 단어장 불러오는 화면으로 돌아가요. (위의 '기록도 함께 지우기' 선택이 그대로 적용돼요)")
        st.checkbox("모든 한자를 지울게요", key="del_all_ok")
        st.button("🗑️ 단어장 전체 삭제", on_click=delete_all_cb, key="del_all_btn", width="stretch",
                  disabled=not s.get("del_all_ok", False))


def render_set_tools():
    with st.expander("📂 내장 단어장 불러오기"):
        set_picker("word", "study", show=False)
    render_word_paste()
    render_word_edit()
    render_word_delete()



def main():
    init_state()
    inject_css()  # 먼저 깔아서, 저장용 보이지 않는 조각이 잠깐이라도 보이지 않게
    if st.query_params.get("nostore"):  # 진단용: 주소 끝에 ?nostore=1 을 붙이면 브라우저 자동 저장을 끄고 실행
        st.session_state.restored, st.session_state.store_ok = True, False
    else:
        sync_browser_store()
    try:
        run_app()
    except Exception:  # 예상 못 한 오류가 나도 하얀 화면 대신 복구 버튼을 보여줘요
        st.error("앗, 화면을 그리다가 문제가 생겼어요. 아래 버튼으로 정리하면 대부분 해결돼요. (단어장과 기록은 그대로예요)")
        st.button("🔄 화면 정리하고 다시 시작", on_click=repair_state, type="primary")
        with st.expander("자세한 오류 내용 (문제가 계속되면 이걸 알려 주세요)"):
            st.code(traceback.format_exc())


main()
