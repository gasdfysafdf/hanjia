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
    set_goal=20, set_stroke=8, set_canvas=280,
    card_src="전체 단어", card_dir="한자 → 뜻·음", card_shuffle=True, card_count=20,
    # 시험 플랜
    set_exam_date="", set_target=90, plan_groups=[],
    # 사자성어
    idm_view="시험", idm_src="전체", idm_type="섞어서", idm_count=10, idm_groups=[],
    # 모의고사
    ex_src="전체 단어", ex_groups=[], ex_apply="반영해요",
    ex_n_ko=10, ex_n_write=5, ex_n_choice=5, ex_n_idiom=0,
    ex_pt_ko=5.0, ex_pt_write=5.0, ex_pt_choice=5.0, ex_pt_idiom=5.0,
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
        notes={}, idioms=[], idiom_stats={}, idiom_wrong={}, exams=[],
        iquiz=[], ipos=0, ires=[], inonce=0, iretry=False, idm_confirm="",
        exam=None, exam_result=None, ex_nonce=0, ex_quit=False,
        ed_ver=0, ied_ver=0, notice=None, hw_clear=0,
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


def reset_all():
    """퀴즈 + 오답 노트 + 통계 + 복습 기록 전부 초기화 (불러온 단어는 유지)"""
    s = st.session_state
    s.wrong, s.stats, s.srs, s.history = {}, {}, {}, {}
    s.idiom_stats, s.idiom_wrong, s.exams = {}, {}, []
    s.cards = None
    s.exam, s.exam_result, s.ex_quit = None, None, False
    reset_idiom_quiz()
    reset_quiz()
    s.confirm_reset = False
    s.flash = "퀴즈, 오답 노트, 통계, 복습·시험 기록을 모두 초기화했어요. (단어장·사자성어·메모는 그대로예요)"


# ----------------------------------------------------------------------
# 새 기능 공통: 상한 · 검증 · 폴더 내장 세트 · 메모/별표 · 단어장 편집
# ----------------------------------------------------------------------
MAX_NOTE = 200  # 메모 한 개의 최대 글자 수
MAX_NOTES = 3000  # 저장하는 메모·별표 개수 (브라우저 저장 용량을 지키려는 상한)
MAX_IDIOMS = 2000  # 사자성어 개수
MAX_EXAMS = 100  # 모의고사 기록 개수
IDIOM_MIN, IDIOM_MAX = 3, 8  # 사자성어 한자 글자 수
MAX_SET_BYTES = 2_000_000  # 폴더 CSV 한 개의 최대 크기

EXAM_ORDER = ["ko", "write", "choice", "idiom"]
EXAM_TYPES = {
    "ko": "한자 → 뜻·음 쓰기",
    "write": "뜻·음 → 한자 쓰기 (손글씨)",
    "choice": "객관식",
    "idiom": "사자성어",
}
EXAM_SRC = ["전체 단어", "분류 선택", "별표 한자만", "오답 노트만"]
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
IDIOM_SRC = ["전체", "분류 선택", "오답 노트만"]
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


def _clean_exam(e):
    if not isinstance(e, dict):
        return None
    day = _iso(e.get("date"))
    if not day:
        return None
    full = _flt(e.get("full"), 100.0, 0.1, 10000.0)
    types = {}
    for k, v in _dict(e.get("types")).items():
        if k in EXAM_TYPES and isinstance(v, dict):
            mx = _flt(v.get("max"), 0.0, 0.0, 10000.0)
            types[k] = {"score": round(_flt(v.get("score"), 0.0, 0.0, mx), 1), "max": round(mx, 1),
                        "n": _int(v.get("n"), 0, 0, 1000), "right": _int(v.get("right"), 0, 0, 1000)}
    return {"start": str(e.get("start", ""))[:19] or day, "date": day,
            "score": round(_flt(e.get("score"), 0.0, 0.0, full), 1), "full": round(full, 1),
            "n": _int(e.get("n"), 0, 0, 1000), "types": types, "secs": _int(e.get("secs"), 0, 0, 10**6)}


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
    """단어장이 바뀔 때: 진행 중이던 퀴즈·카드·시험지·모의고사를 안전하게 비워요 (기록은 그대로)"""
    s = st.session_state
    s.quiz, s.pos, s.results, s.feedback, s.cards, s.sheet = [], 0, [], None, None, None
    s.hint = False
    s.nonce += 1
    s.exam, s.ex_quit = None, False
    s.ex_nonce += 1


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
  var loaded = false, last = null;
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
    "set_goal", "set_stroke", "set_canvas", "set_type", "set_source", "set_count", "set_shuffle", "set_preset",
    "card_src", "card_dir", "card_shuffle", "card_count",
    "set_exam_date", "set_target", "plan_groups", "idm_src", "idm_type", "idm_count",
    "ex_src", "ex_apply",
    "ex_n_ko", "ex_n_write", "ex_n_choice", "ex_n_idiom",
    "ex_pt_ko", "ex_pt_write", "ex_pt_choice", "ex_pt_idiom",
] + list(GRADE_DEFAULTS)
# 저장값 범위 (엉뚱한 값이 들어 있어도 위젯이 깨지지 않게 이 안으로 맞춰요)
SETTING_RANGES = {"set_target": (0, 100), "idm_count": (5, 50)}
SETTING_RANGES.update({f"ex_n_{k}": (0, 100) for k in EXAM_ORDER})
SETTING_RANGES.update({f"ex_pt_{k}": (0.0, 100.0) for k in EXAM_ORDER})
ALLOWED = {
    "set_type": ["섞어서", "뜻·음 쓰기", "한자 쓰기", "객관식"],
    "set_source": SOURCES,
    "card_src": SOURCES,
    "idm_src": IDIOM_SRC,
    "idm_type": IDIOM_KINDS,
    "ex_src": EXAM_SRC,
    "ex_apply": ["반영해요", "반영 안 해요"],
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
    out = {"wrong": {}, "stats": {}, "srs": {}, "history": {}, "notes": {}, "idioms": [],
           "idiom_stats": {}, "idiom_wrong": {}, "exams": [], "plan": {}}
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
    out["idioms"] = _clean_idioms(d.get("idioms"))
    for k, v in _dict(d.get("idiom_stats")).items():
        if isinstance(v, dict) and _clean_idiom_key(k):
            out["idiom_stats"][str(k)] = {"sound": str(v.get("sound", "")), "meaning": str(v.get("meaning", "")),
                                          "right": _int(v.get("right")), "wrong": _int(v.get("wrong"))}
    for k, v in _dict(d.get("idiom_wrong")).items():
        if isinstance(v, dict) and _clean_idiom_key(k) and str(v.get("sound", "")).strip() and str(v.get("meaning", "")).strip():
            out["idiom_wrong"][str(k)] = {
                "hanja": str(k), "sound": str(v.get("sound", "")), "meaning": str(v.get("meaning", "")),
                "wrong_count": _int(v.get("wrong_count"), 1, 1), "last": str(v.get("last", ""))}
    exams = [c for c in (_clean_exam(e) for e in (d.get("exams") if isinstance(d.get("exams"), list) else [])) if c]
    out["exams"] = sorted(exams, key=lambda e: e["start"])[-MAX_EXAMS:]
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
    have = {i["hanja"] for i in s.idioms}  # 사자성어는 이미 있는 것을 유지하고 없는 것만 더해요
    for it in clean.get("idioms", []):
        if it["hanja"] not in have and len(s.idioms) < MAX_IDIOMS:
            s.idioms.append(it)
            have.add(it["hanja"])
    s.idiom_stats.update(clean.get("idiom_stats", {}))
    s.idiom_wrong.update(clean.get("idiom_wrong", {}))
    starts = {e["start"] for e in s.exams}
    for e in clean.get("exams", []):
        if e["start"] not in starts:
            s.exams.append(e)
            starts.add(e["start"])
    s.exams = sorted(s.exams, key=lambda e: e["start"])[-MAX_EXAMS:]
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
         "notes": s.notes, "idioms": s.idioms, "idiom_stats": s.idiom_stats, "idiom_wrong": s.idiom_wrong,
         "exams": s.exams,
         "plan": {"exam_date": s.get("set_exam_date", ""), "target": s.get("set_target", 90),
                  "groups": s.get("plan_groups", [])}},
        ensure_ascii=False, indent=2,
    ).encode("utf-8")


def build_payload():
    s = st.session_state
    data = {"v": 2, "words": s.words, "wrong": s.wrong, "stats": s.stats, "srs": s.srs,
            "history": s.history, "notes": s.notes, "idioms": s.idioms, "idiom_stats": s.idiom_stats,
            "idiom_wrong": s.idiom_wrong, "exams": s.exams,
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
            if k == "set_canvas":  # 엉뚱한 값이 저장돼 있어도 슬라이더가 깨지지 않게 범위 안으로
                v = _int(v, 280, 200, 400)
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
    return bool(s.words or s.wrong or s.stats or s.srs or s.history or s.idioms or s.notes or s.exams
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
    if s.restored and (s.words or s.wrong or s.stats or s.srs or s.history or s.idioms or s.notes or s.exams
                       or s.idiom_stats or s.idiom_wrong):
        save = build_payload()  # 불러오기 전에는 저장하지 않아서, 빈 상태가 기존 기록을 덮어쓰지 않아요
    try:
        with st.container(key="store_box"):
            got = _store_component()(save=save, key="browser_store", default=None)
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
    s.exam, s.exam_result, s.ex_quit = None, None, False
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

/* 모의고사·사자성어 고르기 버튼 (흰 바탕 + 어두운 글씨) */
div[class*="st-key-xc_"] button{min-height:4rem; background:#fff; border:1.5px solid var(--line);}
div[class*="st-key-xc_"] button, div[class*="st-key-xc_"] button p{color:var(--ink);}
div[class*="st-key-xc_text"] button p{font-size:1.1rem;}
div[class*="st-key-xc_hj"] button p, div[class*="st-key-xc_blank"] button p{font-family:var(--serif); font-size:2rem; line-height:1.15;}

/* 탭이 많아서 모바일에서는 조금 좁게 */
@media (max-width:640px){
  .stTabs [data-baseweb="tab"]{padding:0 .6rem; font-size:.92rem; height:2.9rem;}
}

/* 시작 화면 3단계 */
.steps{display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin:.4rem 0 1rem;}
.step{display:flex; flex-direction:column; align-items:center; gap:.4rem; text-align:center; padding:.8rem .4rem;
  background:#fff; border:1px solid var(--line); border-radius:10px; color:var(--ink); font-size:.92rem; line-height:1.35;}
.step b{display:grid; place-items:center; width:30px; height:30px; border-radius:50%; background:var(--indigo);
  color:#fff; font:700 1rem/1 var(--serif);}
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
    f = st.file_uploader("한자 파일 (txt / csv / xlsx)", key="wordfile")  # 형식 제한 없음: 태블릿 호환
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
        if st.button("붙여넣은 내용 적용"):
            words, dropped, merged = build_words(parse_text(pasted))
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
    f = s.get("wordfile")
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
    st.caption("🈴 사자성어와 모의고사는 단어장을 불러온 뒤 '사자성어'·'시험' 탭에서 쓸 수 있어요.")
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
    return _int(st.session_state.get("set_canvas"), 280, 200, 400)


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
    left, right = st.columns([1, 1])
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
        st.markdown("##### 최근 14일")
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
        st.dataframe(df, width="stretch", hide_index=True, height=480)
    else:
        cards = []
        for w in words:
            mean = f"<b>{highlight(w['meaning'], terms)}</b> {highlight(w['sound'], terms)}"
            body = f"<details><summary>보기</summary>{mean}</details>" if hide else mean
            g = f"<div class='g'>{esc(w['group'])}</div>" if w.get("group") else ""
            star = "<span class='star'>★</span>" if is_star(w["hanja"]) else ""
            nt = note_of(w["hanja"])
            note_html = f"<div class='note' title='{esc(nt)}'>📝 {esc(nt)}</div>" if nt else ""
            cards.append(f"<div class='wcard'>{star}<div class='h'>{highlight(w['hanja'], terms)}</div><div class='m'>{body}</div>{g}{note_html}</div>")
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
    st.sidebar.slider("쓰는 칸 크기", 200, 400, step=20, key="set_canvas",
                      help="한자를 쓰는 칸의 크기예요. 화면이 작으면 줄이세요. (바꾸면 쓰던 글씨는 지워져요)")
    grading_sidebar()

    with st.sidebar.expander("💾 진행 상황 저장 / 불러오기"):
        st.caption("평소엔 이 기기에 자동 저장돼요. 다른 기기로 옮기거나 백업할 때만 쓰세요. (오답 노트·통계·복습 기록·날짜별 기록·메모·사자성어·모의고사 기록·플랜 설정)")
        st.download_button("💾 저장", progress_json(), "hanja_progress.json", "application/json", width="stretch",
                           key="backup_side")
        pf = st.file_uploader("불러오기", type=["json"], key="progress_file")
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
        st.checkbox("전체 초기화 확인 (오답 노트·통계·복습·시험 기록 삭제)", key="confirm_reset")
        st.button("🗑️ 전체 초기화", on_click=reset_all, disabled=not s.get("confirm_reset", False), width="stretch")


def run_app():
    sidebar()
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
    t1, t2, t3, t4, t5, t6, t7 = st.tabs(
        ["✏️ 연습", "🃏 카드", "💯 시험", "📜 사자성어", "📕 오답", "📊 기록", "📚 단어장"])
    with t1:
        render_practice()
    with t2:
        render_cards()
    with t3:
        render_exam()
    with t4:
        render_idioms()
    with t5:
        render_wrong()
    with t6:
        render_stats()
    with t7:
        render_study()


# ----------------------------------------------------------------------
# 사자성어 (관리 · 시험 · 오답 노트) — 한자 통계와 섞지 않고 따로 저장해요
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
    """보기 만들기: 같은 분류·같은 글자 수의 다른 사자성어를 먼저 오답 보기로 써요"""
    s = st.session_state
    key = (lambda x: norm(x["meaning"])) if field == "meaning" else (lambda x: x["hanja"])
    others = [x for x in s.idioms if x["hanja"] != it["hanja"]]
    if avoid_sound:  # 독음이 같은 다른 사자성어는 정답 시비가 생겨서 제외
        others = [x for x in others if norm(x["sound"]) != norm(it["sound"])]
    random.shuffle(others)
    grp, ln = it.get("group"), len(it["hanja"])
    others.sort(key=lambda x: (0 if grp and x.get("group") == grp else 1, 0 if len(x["hanja"]) == ln else 1))
    seen, pick = {key(it)}, []
    for x in others:
        if key(x) in seen:
            continue
        seen.add(key(x))
        pick.append(x)
        if len(pick) == n - 1:
            break
    opts = pick + [it]
    random.shuffle(opts)
    return [{"key": x["hanja"], "text": x[field]} for x in opts]


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
    n = len(st.session_state.idioms)
    choice_ok = n >= 4
    blank_ok = choice_ok and len(idiom_char_pool()) >= 10
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


def record_external(w, qtype, ok, ans, retry=False):
    """기존 record()를 재사용하되, 연습 중인 퀴즈의 결과 목록(s.results)은 건드리지 않게 감싸요 (모의고사용)"""
    s = st.session_state
    keep, keep_retry = s.results, s.retry_mode
    s.results, s.retry_mode = [], bool(retry)
    try:
        record(w, qtype, ok, ans)
    finally:
        s.results, s.retry_mode = keep, keep_retry


# --- 손글씨 공통 (모의고사·사자성어 빈칸용; 연습 탭의 render_write는 그대로 둬요) -------
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


# --- 문제 그리기 (사자성어 시험 · 모의고사 공용) -------------------------------------
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
        left, right = st.columns(2)
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
def idiom_pool_for(src, groups):
    s = st.session_state
    if src == "오답 노트만":
        return [{"hanja": v["hanja"], "sound": v["sound"], "meaning": v["meaning"], "group": ""}
                for v in s.idiom_wrong.values() if v.get("sound") and v.get("meaning")]
    return [dict(x) for x in s.idioms if not (src == "분류 선택" and groups) or x.get("group") in groups]


def start_idiom_quiz():
    s = st.session_state
    src, kind = s.get("idm_src", "전체"), s.get("idm_type", "섞어서")
    pool = idiom_pool_for(src, s.get("idm_groups", []))
    if not pool:
        s.flash = ("오답 노트에 사자성어가 없어요. 🎉" if src == "오답 노트만"
                   else "문제로 만들 사자성어가 없어요. '관리'에서 먼저 추가해 주세요.")
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


def go_idiom_manage():
    st.session_state.idm_view = "관리"


def render_idiom_setup():
    s = st.session_state
    n = len(s.idioms)
    st.subheader("사자성어 시험")
    st.caption(f"사자성어 {n}개 · 한자 시험과 따로 기록해요. 정답은 끝난 뒤에 보여줘요.")
    av = idiom_kinds_avail()
    kinds = [k for k in IDIOM_KINDS if idiom_types_for(k, av)]
    if n < 4:
        st.info("사자성어가 4개 미만이라 '고르기·빈칸' 유형은 쓸 수 없어요. 4개 이상 담으면 열려요. (지금은 독음 쓰기만 가능해요)")
    elif "빈칸 손글씨" not in kinds:
        st.caption("'빈칸 손글씨'는 손글씨 기능과 한자 폰트가 있어야 해요. (지금은 빠져 있어요)")
    if s.get("idm_type") not in kinds:
        s.idm_type = kinds[0]
    st.radio("유형", kinds, key="idm_type", horizontal=True)
    if s.get("idm_src") not in IDIOM_SRC:
        s.idm_src = IDIOM_SRC[0]
    st.radio("범위", IDIOM_SRC, key="idm_src", horizontal=True)
    if s.idm_src == "분류 선택":
        groups = sorted({x["group"] for x in s.idioms if x.get("group")})
        if groups:
            s.idm_groups = [g for g in s.get("idm_groups", []) if g in groups]
            st.multiselect("분류 (비워 두면 전체)", groups, key="idm_groups", placeholder="전체 (눌러서 고르기)")
        else:
            st.caption("분류가 적힌 사자성어가 없어서 전체에서 나와요.")
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
        st.info("아직 사자성어가 없어요. '관리'에서 추가하거나 내장 세트를 불러와 주세요.")
        st.button("🛠 관리로 가기", on_click=go_idiom_manage, key="idm_go_manage")
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


# --- 사자성어 관리 ---------------------------------------------------------------------
IDIOM_COLS = ["사자성어", "독음", "뜻", "분류"]


def idioms_df(lst):
    return pd.DataFrame([{"사자성어": x["hanja"], "독음": x["sound"], "뜻": x["meaning"], "분류": x.get("group", "")}
                         for x in lst], columns=IDIOM_COLS)


def idioms_csv(lst):
    return idioms_df(lst).to_csv(index=False).encode("utf-8-sig")


def idioms_from_df(df):
    out, errors, seen = [], [], {}
    for i, (_, r) in enumerate(df.iterrows(), 1):
        h = re.sub(r"\s+", "", _cell(r.get("사자성어")))
        sd, mn, g = _cell(r.get("독음")), _cell(r.get("뜻")), _cell(r.get("분류"))
        if not (h or sd or mn or g):
            continue
        why = idiom_problem(h, sd, mn)
        if why:
            errors.append(f"{i}번째 줄 ({h or '한자 없음'}): {why}")
            continue
        if h in seen:
            errors.append(f"{i}번째 줄 ({h}): {seen[h]}번째 줄과 같은 사자성어예요")
            continue
        seen[h] = i
        out.append({"hanja": h, "sound": sd, "meaning": mn, "group": g[:30]})
    if len(out) > MAX_IDIOMS:
        errors.append(f"사자성어는 최대 {MAX_IDIOMS}개까지만 저장해요 (지금 {len(out)}개)")
    return out, errors


def add_idioms(items):
    """새 사자성어 합치기 (이미 있는 것은 먼저 있던 걸 유지). 반환: (새로 담은 수, 겹친 수, 뜻·음이 달라 유지한 것들)"""
    s = st.session_state
    merged, added, dup, conf = merge_items(s.idioms, items, ("meaning", "sound"))
    if len(merged) > MAX_IDIOMS:
        merged = merged[:MAX_IDIOMS]
    s.idioms = merged
    reset_idiom_quiz()
    s.ied_ver += 1
    return added, dup, conf


def _report_add(items, probs, dup_in):
    """붙여넣기·파일로 가져온 결과를 합치고 안내 문구를 만들어요 (콜백/버튼 공용)"""
    s = st.session_state
    notes = []
    if items:
        added, dup, conf = add_idioms(items)
        msg = f"사자성어 {added}개를 담았어요."
        if dup + dup_in:
            msg += f" (겹쳐서 건너뜀 {dup + dup_in}개)"
        if conf:
            notes.append(f"같은 사자성어인데 독음·뜻이 달라서 먼저 있던 것을 유지한 항목 {len(conf)}개: " + ", ".join(conf[:10]))
        s.flash = msg
    else:
        s.flash = "담을 수 있는 사자성어가 없었어요."
    if probs:
        notes.append(f"제외한 줄 {len(probs)}개:\n" + "\n".join(f"- {p}" for p in probs[:8])
                     + (f"\n- … 외 {len(probs) - 8}건" if len(probs) > 8 else ""))
    s.notice = {"kind": "idiom", "msgs": notes} if notes else None


def import_idiom_paste():
    s = st.session_state
    ver = s.ied_ver
    text, grp = s.get(f"idp_t_{ver}", ""), s.get(f"idp_g_{ver}", "")
    items, probs, dup = build_idioms(parse_idiom_text(text or "", (grp or "").strip()[:30]))
    _report_add(items, probs, dup)


def import_idiom_file():
    s = st.session_state
    f = s.get(f"idiom_file_{s.ied_ver}")
    if f is None:
        s.flash = "먼저 파일을 골라 주세요."
        return
    if f.size > MAX_SET_BYTES * 5:
        s.flash = "파일이 너무 커요 (10MB 이하)."
        return
    try:
        rows = parse_idiom_file(f)
    except Exception as e:
        s.flash = f"파일을 읽지 못했어요: {e}"
        return
    items, probs, dup = build_idioms(rows)
    _report_add(items, probs, dup)


def commit_idioms(new):
    """표 편집 결과 저장. 독음·뜻을 고치면 오답 노트·통계의 복사본도 같이 바꿔요 (한자를 바꾸면 삭제 + 새 항목)"""
    s = st.session_state
    old = {x["hanja"]: x for x in s.idioms}
    keys = {x["hanja"] for x in new}
    changed = 0
    for x in new:
        o = old.get(x["hanja"])
        if o and (o["sound"] != x["sound"] or o["meaning"] != x["meaning"]):
            changed += 1
            for book in (s.idiom_wrong, s.idiom_stats):
                e = book.get(x["hanja"])
                if e:
                    e["sound"], e["meaning"] = x["sound"], x["meaning"]
    added = sum(1 for x in new if x["hanja"] not in old)
    removed = sum(1 for h in old if h not in keys)
    s.idioms = new
    reset_idiom_quiz()
    s.ied_ver += 1
    s.flash = f"사자성어를 저장했어요. (추가 {added} · 수정 {changed} · 삭제 {removed})"


def idiom_confirm(kind):
    st.session_state.idm_confirm = kind


def idiom_delete_do():
    s = st.session_state
    if s.get("idm_confirm") == "all":
        n = len(s.idioms)
        s.idioms = []
    else:
        sel = set(s.get("idm_del", []))
        n = len(sel)
        s.idioms = [x for x in s.idioms if x["hanja"] not in sel]
    s.idm_del, s.idm_confirm = [], ""
    reset_idiom_quiz()
    s.ied_ver += 1
    s.flash = f"사자성어 {n}개를 지웠어요. (풀이 기록은 남겨 뒀어요)"


def render_idiom_manage():
    s = st.session_state
    ver = s.ied_ver
    show_notice("idiom", "idm_top")
    st.caption(f"사자성어 {len(s.idioms)}개 (최대 {MAX_IDIOMS}개) · 한 사자성어는 한자 {IDIOM_MIN}~{IDIOM_MAX}자, 독음·뜻은 꼭 있어야 해요.")
    with st.expander("📂 내장 사자성어 불러오기"):
        set_picker("idiom", "idiom", show=False)
    with st.expander("➕ 한 개 추가"):
        with st.form("idm_add_form", border=False):
            c1, c2 = st.columns(2)
            h = c1.text_input("사자성어 (한자)", key=f"ida_h_{ver}")
            sd = c2.text_input("독음", key=f"ida_s_{ver}")
            m = st.text_input("뜻", key=f"ida_m_{ver}")
            g = st.text_input("분류 (선택)", key=f"ida_g_{ver}")
            add = st.form_submit_button("➕ 사자성어 추가", type="primary", width="stretch")
        if add:
            h, sd, m, g = re.sub(r"\s+", "", h or ""), (sd or "").strip(), (m or "").strip(), (g or "").strip()[:30]
            why = idiom_problem(h, sd, m)
            if not why and any(x["hanja"] == h for x in s.idioms):
                why = "이미 있는 사자성어예요. '표로 고치기'에서 수정해 주세요."
            if not why and len(s.idioms) >= MAX_IDIOMS:
                why = f"사자성어는 최대 {MAX_IDIOMS}개까지만 담아요."
            if why:
                st.warning(why)
            else:
                add_idioms([{"hanja": h, "sound": sd, "meaning": m, "group": g}])
                s.flash = f"'{h}'을(를) 추가했어요."
                st.rerun()
    with st.expander("📋 여러 줄 붙여넣기"):
        with st.form("idm_paste_form", border=False):
            st.text_area("한 줄에 하나씩 (한자 독음 뜻)", height=150, key=f"idp_t_{ver}",
                         placeholder="一石二鳥 일석이조 한 가지 일로 두 가지 이익을 얻음")
            st.text_input("분류 (선택, 붙여넣은 줄 전부에 같이 적어요)", key=f"idp_g_{ver}")
            st.form_submit_button("📋 붙여넣은 줄 추가", type="primary", on_click=import_idiom_paste, width="stretch")
        st.caption("공백으로 나눠도 되고, 탭·'|'·';'·쉼표(한자,독음,뜻)로 나눠도 돼요. '#'으로 시작하는 줄은 무시해요.")
    with st.expander("📤 CSV 올리기 · 받기"):
        st.file_uploader("사자성어 파일 (csv / xlsx / txt)", key=f"idiom_file_{ver}")  # 형식 제한 없음: 태블릿 호환
        st.button("📤 올린 파일 불러오기 (지금 목록에 합치기)", on_click=import_idiom_file, width="stretch", key="idm_import")
        st.caption("열: 사자성어, 독음, 뜻, 분류(선택). 첫 줄은 제목 줄이어도 돼요.")
        st.download_button("📥 현재 사자성어 CSV 받기", idioms_csv(s.idioms), "내_사자성어.csv", "text/csv",
                           width="stretch", key="idm_csv_dl")
        st.caption("받은 파일을 GitHub 저장소의 idioms/ 폴더에 올리면 '내장 사자성어'가 돼요. 파일 이름이 세트 이름이에요.")
    with st.expander("✏️ 표로 고치기 · 삭제"):
        if s.idioms:
            with st.form("idm_edit_form", border=False):
                edited = st.data_editor(
                    idioms_df(s.idioms), num_rows="dynamic", hide_index=True, key=f"idm_editor_{ver}",
                    column_config={
                        "사자성어": st.column_config.TextColumn("사자성어", max_chars=IDIOM_MAX),
                        "독음": st.column_config.TextColumn("독음", max_chars=60),
                        "뜻": st.column_config.TextColumn("뜻", max_chars=300),
                        "분류": st.column_config.TextColumn("분류", max_chars=30),
                    })
                saved = st.form_submit_button("💾 사자성어 저장", type="primary", width="stretch")
            if saved:
                new, errs = idioms_from_df(edited)
                if errs:
                    st.error("저장하지 못했어요. 아래를 고친 뒤 다시 눌러 주세요.\n\n"
                             + "\n".join(f"- {e}" for e in errs[:8])
                             + (f"\n- … 외 {len(errs) - 8}건" if len(errs) > 8 else ""))
                elif new == s.idioms:
                    st.info("바뀐 내용이 없어요.")
                else:
                    commit_idioms(new)
                    st.rerun()
            st.markdown("**골라서 삭제**")
            keys = [x["hanja"] for x in s.idioms]
            snd = {x["hanja"]: x["sound"] for x in s.idioms}
            s.idm_del = [x for x in s.get("idm_del", []) if x in snd]
            st.multiselect("지울 사자성어", keys, key="idm_del", format_func=lambda k: f"{k} {snd.get(k, '')}",
                           placeholder="여러 개 고르기")
            if s.get("idm_confirm"):
                n = len(s.idioms) if s.idm_confirm == "all" else len(s.idm_del)
                st.warning(f"정말 {n}개를 지울까요? 되돌릴 수 없어요.")
                y, no = st.columns(2)
                y.button("네, 지울게요", on_click=idiom_delete_do, width="stretch", key="idm_del_yes")
                no.button("아니요", on_click=idiom_confirm, args=("",), width="stretch", key="idm_del_no")
            else:
                d1, d2 = st.columns(2)
                d1.button("선택 삭제", on_click=idiom_confirm, args=("del",), disabled=not s.idm_del,
                          width="stretch", key="idm_del_ask")
                d2.button("전체 비우기", on_click=idiom_confirm, args=("all",), width="stretch", key="idm_all_ask")
        else:
            st.caption("아직 사자성어가 없어요. 위에서 추가하거나 불러와 주세요.")


def render_idioms():
    s = st.session_state
    views = ["시험", "관리", "오답 노트"]
    if s.get("idm_view") not in views:
        s.idm_view = views[0]
    st.radio("사자성어 화면", views, key="idm_view", horizontal=True, label_visibility="collapsed")
    if s.idm_view == "관리":
        render_idiom_manage()
    elif s.idm_view == "오답 노트":
        render_idiom_wrong()
    else:
        render_idiom_exam()


# ----------------------------------------------------------------------
# 모의고사 (100점 만점) — 정답은 끝난 뒤에 한꺼번에 보여줘요
# ----------------------------------------------------------------------
def exam_pool():
    """모의고사 출제 범위에 해당하는 한자들"""
    s = st.session_state
    src = s.get("ex_src", "전체 단어")
    if src == "오답 노트만":
        return [{"hanja": v["hanja"], "meaning": v["meaning"], "sound": v["sound"], "group": ""}
                for v in s.wrong.values() if v.get("sound")]
    groups = s.get("ex_groups", []) if src == "분류 선택" else []
    pool = [w for w in s.words if not groups or w.get("group") in groups]
    if src == "별표 한자만":
        pool = [w for w in pool if is_star(w["hanja"])]
    return pool


def exam_idiom_pool():
    s = st.session_state
    if s.get("ex_src") == "오답 노트만":
        return idiom_pool_for("오답 노트만", [])
    return [dict(x) for x in s.idioms]


def exam_limits(pool):
    """유형별로 지금 낼 수 있는 최대 문항 수 (0이면 막힌 유형)"""
    s = st.session_state
    font_ok = bool(HAS_CANVAS and find_font())
    return {
        "ko": len(pool),
        "write": sum(1 for w in pool if len(w["hanja"]) == 1) if font_ok else 0,
        "choice": len(pool) if len(s.words) >= 4 else 0,
        "idiom": len(exam_idiom_pool()) if len(s.idioms) >= 4 else 0,
    }


def exam_blocked_reasons(lim, pool):
    s = st.session_state
    r = {}
    if not pool:
        r["ko"] = "고른 범위에 한자가 없어요"
    if lim["write"] == 0:
        r["write"] = ("손글씨 기능(캔버스)이나 한자 폰트가 없어요" if not (HAS_CANVAS and find_font())
                      else "범위 안에 한 글자 한자가 없어요")
    if lim["choice"] == 0:
        r["choice"] = "단어장에 한자가 4개 이상 있고 범위 안에 한자가 있어야 해요"
    if lim["idiom"] == 0:
        r["idiom"] = (f"사자성어가 4개 이상 있어야 해요 (지금 {len(s.idioms)}개)" if len(s.idioms) < 4
                      else "범위(오답 노트)에 사자성어가 없어요")
    return r


def exam_total(lim=None):
    """(총 문항 수, 배점 합계) — 막힌 유형은 0문항으로 쳐요"""
    s = st.session_state
    lim = lim if lim is not None else exam_limits(exam_pool())
    n_all, tot = 0, 0.0
    for k in EXAM_ORDER:
        n = 0 if lim[k] == 0 else _int(s.get(f"ex_n_{k}"), 0, 0, 100)
        n_all += n
        tot += n * _flt(s.get(f"ex_pt_{k}"), 0.0, 0.0, 100.0)
    return n_all, round(tot, 2)


def fit_points():
    """배점을 비례해서 합계 100점으로 맞춰요 (소수 첫째 자리)"""
    s = st.session_state
    lim = exam_limits(exam_pool())
    n = {k: (0 if lim[k] == 0 else _int(s.get(f"ex_n_{k}"), 0, 0, 100)) for k in EXAM_ORDER}
    act = [k for k in EXAM_ORDER if n[k] > 0]
    if not act:
        return
    pt = {k: _flt(s.get(f"ex_pt_{k}"), 0.0, 0.0, 100.0) for k in EXAM_ORDER}
    total = sum(n[k] * pt[k] for k in act)
    if total <= 0:
        eq = 100.0 / sum(n[k] for k in act)
        pt.update({k: eq for k in act})
        total = 100.0
    f = 100.0 / total
    for k in act:
        pt[k] = round(min(100.0, pt[k] * f), 1)
    resid = round(100.0 - sum(n[k] * pt[k] for k in act), 4)
    if abs(resid) > 1e-6:  # 반올림으로 생긴 차이를 한 유형에서 딱 맞게 메울 수 있으면 메워요
        for k in reversed(act):
            fixed = round(pt[k] + resid / n[k], 1)
            if 0 <= fixed <= 100 and abs(sum(n[j] * (fixed if j == k else pt[j]) for j in act) - 100.0) < 1e-6:
                pt[k] = fixed
                break
    for k in act:
        s[f"ex_pt_{k}"] = float(pt[k])
    s.flash = "배점을 100점에 맞췄어요."


def start_exam():
    s = st.session_state
    pool = exam_pool()
    lim = exam_limits(pool)
    want = {k: (0 if lim[k] == 0 else _int(s.get(f"ex_n_{k}"), 0, 0, 100)) for k in EXAM_ORDER}
    pts = {k: _flt(s.get(f"ex_pt_{k}"), 0.0, 0.0, 100.0) for k in EXAM_ORDER}
    if sum(want.values()) == 0:
        s.flash = "문항이 0개예요. 문항 수를 정해 주세요."
        return
    random.shuffle(pool)
    singles = [w for w in pool if len(w["hanja"]) == 1]
    write_w = singles[: want["write"]]  # 같은 한자가 여러 유형에 겹쳐 나오지 않게 나눠 뽑아요
    used = {w["hanja"] for w in write_w}
    rest = [w for w in pool if w["hanja"] not in used]
    ko_w, rest = rest[: want["ko"]], rest[want["ko"]:]
    ch_w = rest[: want["choice"]]
    items = [{"kind": "ko", "pts": pts["ko"], "q": {"w": w, "type": "read"}} for w in ko_w]
    items += [{"kind": "write", "pts": pts["write"], "q": {"w": w, "type": "write"}} for w in write_w]
    items += [{"kind": "choice", "pts": pts["choice"],
               "q": make_item(w, random.choice(["choice_ko", "choice_hj"]))} for w in ch_w]
    got = {"ko": len(ko_w), "write": len(write_w), "choice": len(ch_w), "idiom": 0}
    if want["idiom"]:
        ipool = exam_idiom_pool()
        random.shuffle(ipool)
        types = [t for t in idiom_types_for("섞어서") if t != "i_hw"]
        cpool = idiom_char_pool()
        picked = ipool[: want["idiom"]]
        items += [{"kind": "idiom", "pts": pts["idiom"], "q": make_idiom_item(it, random.choice(types), cpool)}
                  for it in picked]
        got["idiom"] = len(picked)
    if not items:
        s.flash = "문제를 만들 수 없어요. 범위를 확인해 주세요."
        return
    short = [f"{EXAM_TYPES[k]} {want[k]}→{got[k]}문항" for k in EXAM_ORDER if want[k] and got[k] < want[k]]
    s.exam = {"items": items, "pos": 0, "answers": [], "start": now_kst_str(), "t0": time.time(),
              "apply": s.get("ex_apply", "반영해요") == "반영해요", "src": s.get("ex_src", "전체 단어")}
    s.exam_result, s.ex_quit = None, False
    s.ex_nonce += 1
    s.flash = ("범위 안 한자가 모자라서 문항을 줄였어요: " + ", ".join(short)) if short else ""


def exam_wrong_row(n, it, a):
    k, q = it["kind"], it["q"]
    if k == "idiom":
        x = q["it"]
        return {"번호": n, "유형": f"{EXAM_TYPES[k]} · {IDIOM_LABELS[q['type']]}", "문제": idiom_q_text(q),
                "내 답": a["answer"], "정답": f"{idiom_a_text(q)}  ({x['hanja']} · {x['meaning']})", "내 메모": ""}
    w, ty = q["w"], q["type"]
    if k == "write" or ty == "choice_hj":
        question, right = _label(w), w["hanja"]
    else:
        question, right = w["hanja"], _label(w)
    return {"번호": n, "유형": EXAM_TYPES[k], "문제": question, "내 답": a["answer"], "정답": right,
            "내 메모": note_of(w["hanja"])}


def finish_exam():
    """채점 결과를 정리해서 기록(s.exams)에 남기고, 옵션에 따라 오답 노트·복습 기록에 반영해요"""
    s = st.session_state
    ex = s.exam
    if not ex:
        return
    secs = int(max(0, time.time() - ex["t0"]))
    types = {k: {"score": 0.0, "max": 0.0, "n": 0, "right": 0} for k in EXAM_ORDER}
    wrong = []
    for n, (it, a) in enumerate(zip(ex["items"], ex["answers"]), 1):
        t = types[it["kind"]]
        t["max"] += it["pts"]
        t["n"] += 1
        if a["ok"]:
            t["score"] += it["pts"]
            t["right"] += 1
        else:
            wrong.append(exam_wrong_row(n, it, a))
    types = {k: {"score": round(v["score"], 1), "max": round(v["max"], 1), "n": v["n"], "right": v["right"]}
             for k, v in types.items() if v["n"]}
    full = round(sum(v["max"] for v in types.values()), 1)
    score = round(sum(v["score"] for v in types.values()), 1)
    entry = {"start": ex["start"], "date": ex["start"][:10], "score": score, "full": full or 100.0,
             "n": len(ex["items"]), "types": types, "secs": secs}
    s.exams.append(entry)
    s.exams = s.exams[-MAX_EXAMS:]
    if ex["apply"]:
        retry = ex.get("src") == "오답 노트만"
        for it, a in zip(ex["items"], ex["answers"]):
            q = it["q"]
            if it["kind"] == "idiom":
                record_idiom(q, a["ok"], a["answer"], retry=retry)
            else:
                record_external(q["w"], q["type"], a["ok"], a["answer"], retry=retry)
    s.exam_result = {"entry": entry, "wrong": wrong, "applied": bool(ex["apply"]),
                     "end": now_kst_str()}
    s.exam, s.ex_quit = None, False


def exam_answer(ok, shown):
    s = st.session_state
    ex = s.exam
    if not ex or ex["pos"] >= len(ex["items"]):
        return
    ex["answers"].append({"ok": bool(ok), "answer": str(shown)})
    ex["pos"] += 1
    s.ex_nonce += 1
    if ex["pos"] >= len(ex["items"]):
        finish_exam()


def exam_idiom_answer(ok, shown):
    exam_answer(ok, shown)


def exam_pick_choice(q, o):
    w, ty = q["w"], q["type"]
    if o is None:
        exam_answer(False, "(모르겠어요)")
        return
    shown = o["hanja"] if ty == "choice_hj" else o["label"]
    exam_answer(o["hanja"] == w["hanja"], shown)


def exam_quit_ask():
    st.session_state.ex_quit = True


def exam_quit_cancel():
    st.session_state.ex_quit = False


def exam_quit():
    s = st.session_state
    s.exam, s.ex_quit = None, False
    s.ex_nonce += 1
    s.flash = "시험을 그만뒀어요. 기록에는 남기지 않았어요."


def close_exam_result():
    st.session_state.exam_result = None


# --- 화면: 시험 보는 중 ---------------------------------------------------------------
def render_exam_ko(q):
    s = st.session_state
    w = q["w"]
    hanja_card(w["hanja"])
    html("<div class='ask'>이 한자의 <b>뜻과 음</b>을 입력하세요. 예) 하늘 천</div>")
    base = f"{s.exam['pos']}_{s.ex_nonce}"
    with st.form(f"exf_{base}", border=False):
        ans = st.text_input("답", key=f"exa_{base}", label_visibility="collapsed", placeholder="뜻과 음 (예: 하늘 천)")
        c1, c2 = st.columns(2)
        go = c1.form_submit_button("다음 문제 ➡️", type="primary", width="stretch")
        skip = c2.form_submit_button("모르겠어요", width="stretch")
    if go or skip:
        a = (ans or "").strip() if go else ""
        ok = check_reading(a, w)[0] if a else False
        exam_answer(ok, a or ("(빈칸)" if go else "(모르겠어요)"))
        st.rerun()


def render_exam_write(q):
    s = st.session_state
    w = q["w"]
    font, px = find_font(), canvas_px()
    base = f"{s.exam['pos']}_{s.ex_nonce}"
    left, right = st.columns(2)
    with left:
        meaning_card(w)
        html("<div class='ask'>에 해당하는 한자를 오른쪽 칸에 쓰세요</div>")
    canvas_backdrop_for("canvas_box_exam", grid_image(px), px)
    with right:
        with st.container(key="canvas_box_exam"):
            res = draw_canvas(f"xcv_{base}_{s.hw_clear}_{px}", px)
    b1, b2, b3 = st.columns(3)
    if b1.button("다음 문제 ➡️", key=f"xw_go_{base}", type="primary", width="stretch"):
        cands = [x["hanja"] for x in s.words if len(x["hanja"]) == 1]
        ok, shown = hw_grade(res, w["hanja"], font, cands)
        exam_answer(ok, shown)
        st.rerun()
    b2.button("지우기", key=f"xw_clr_{base}", on_click=hw_clear, width="stretch")
    if b3.button("모르겠어요", key=f"xw_skip_{base}", width="stretch"):
        exam_answer(False, "(모르겠어요)")
        st.rerun()


def render_exam_choice(q):
    s = st.session_state
    w, ty = q["w"], q["type"]
    if ty == "choice_ko":
        hanja_card(w["hanja"])
        html("<div class='ask'>이 한자의 <b>뜻과 음</b>은?</div>")
    else:
        meaning_card(w)
        html("<div class='ask'>이 뜻·음에 해당하는 <b>한자</b>는?</div>")
    base = f"{s.exam['pos']}_{s.ex_nonce}"
    with st.container(key="xc_text_exam" if ty == "choice_ko" else "xc_hj_exam"):
        cols = st.columns(2)
        for i, o in enumerate(q["opts"]):
            text = o["label"] if ty == "choice_ko" else o["hanja"]
            cols[i % 2].button(text, key=f"xo_{base}_{i}", on_click=exam_pick_choice, args=(q, o), width="stretch")
    st.button("모르겠어요", key=f"xo_skip_{base}", on_click=exam_pick_choice, args=(q, None))


def render_exam_run():
    s = st.session_state
    ex = s.exam
    items, pos = ex["items"], ex["pos"]
    it = items[pos]
    st.progress(pos / len(items), text=f"{pos + 1} / {len(items)}  ·  {EXAM_TYPES[it['kind']]}  ·  {fmt_pts(it['pts'])}점")
    if s.ex_quit:
        st.warning("정말 그만둘까요? 지금까지 푼 답은 기록에 남지 않아요.")
        c1, c2 = st.columns(2)
        c1.button("네, 그만둘래요", key="ex_quit_yes", on_click=exam_quit, width="stretch")
        c2.button("계속 풀게요", key="ex_quit_no", on_click=exam_quit_cancel, width="stretch", type="primary")
        return
    kind = it["kind"]
    if kind == "ko":
        render_exam_ko(it["q"])
    elif kind == "write":
        render_exam_write(it["q"])
    elif kind == "choice":
        render_exam_choice(it["q"])
    else:
        render_idiom_q(it["q"], "exm", exam_idiom_answer, pos, s.ex_nonce)
    st.button("그만두기", key="ex_quit_btn", on_click=exam_quit_ask)


# --- 화면: 기록 그래프 / 결과 -----------------------------------------------------------
def exam_pct(e):
    return round(e["score"] / e["full"] * 100, 1) if e.get("full") else 0.0


def render_exam_history(limit=20):
    s = st.session_state
    if not s.exams:
        st.caption("아직 본 모의고사가 없어요.")
        return
    last = s.exams[-limit:]
    target = _int(s.get("set_target"), 90, 0, 100)
    first = len(s.exams) - len(last) + 1
    df = pd.DataFrame({"회차": list(range(first, first + len(last))),
                       "내 점수": [exam_pct(e) for e in last], "목표 점수": [target] * len(last)})
    st.line_chart(df, x="회차", y=["내 점수", "목표 점수"], color=["#2B3A67", "#B8372C"])
    st.caption("점수는 100점 만점으로 환산한 값이에요. 빨간 선이 목표 점수예요.")
    st.dataframe(pd.DataFrame([{"회차": first + i, "날짜": e["date"], "점수": f"{fmt_pts(e['score'])} / {fmt_pts(e['full'])}",
                                "문항": e["n"], "걸린 시간": fmt_secs(e["secs"])} for i, e in enumerate(last)][::-1]),
                 width="stretch", hide_index=True)


def render_exam_result():
    s = st.session_state
    r = s.exam_result
    e = r["entry"]
    st.subheader("모의고사 결과예요")
    html(f"<div class='score-box'><div class='big'>{fmt_pts(e['score'])} / {fmt_pts(e['full'])}점</div>"
         f"<div class='lbl'>{e['n']}문항 · 걸린 시간 {fmt_secs(e['secs'])}</div></div>")
    chips = []
    if len(s.exams) >= 2:
        diff = round(exam_pct(e) - exam_pct(s.exams[-2]), 1)
        chips.append(f"<span class='chip {'good' if diff >= 0 else 'hot'}'>지난 시험보다 {diff:+g}점</span>")
    target = _int(s.get("set_target"), 90, 0, 100)
    chips.append(f"<span class='chip {'good' if exam_pct(e) >= target else ''}'>목표 {target}점</span>")
    chips.append("<span class='chip'>" + ("오답 노트·복습 기록에 반영했어요" if r["applied"]
                                       else "오답 노트·복습 기록에는 반영하지 않았어요") + "</span>")
    html("<div class='chips'>" + "".join(chips) + "</div>")
    st.markdown("##### 유형별 점수")
    st.dataframe(pd.DataFrame([{"유형": EXAM_TYPES[k], "맞힌 문항": f"{v['right']} / {v['n']}",
                                "점수": f"{fmt_pts(v['score'])} / {fmt_pts(v['max'])}"}
                               for k, v in e["types"].items()]), width="stretch", hide_index=True)
    if r["wrong"]:
        st.markdown("##### 틀린 문제")
        st.dataframe(pd.DataFrame(r["wrong"]), width="stretch", hide_index=True)
    else:
        st.success("전부 맞혔어요! 🎉")
    st.markdown("##### 점수 추이")
    render_exam_history()
    st.button("🔁 새 모의고사 보기", type="primary", on_click=close_exam_result, width="stretch", key="ex_new")


# --- 화면: 시작 설정 ------------------------------------------------------------------
def render_exam_setup():
    s = st.session_state
    st.subheader("모의고사 (100점 만점)")
    st.caption("정답은 끝난 뒤에 한꺼번에 보여줘요. 힌트는 없고, 시간 제한도 없어요. 걸린 시간만 기록해요. "
               "(모르겠어요는 오답, 빈칸도 괜찮아요)")
    if s.get("ex_src") not in EXAM_SRC:
        s.ex_src = EXAM_SRC[0]
    st.radio("출제 범위", EXAM_SRC, key="ex_src", horizontal=True)
    if s.ex_src == "분류 선택":
        groups = sorted({w["group"] for w in s.words if w.get("group")},
                        key=lambda g: [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", g)])
        if groups:
            s.ex_groups = [g for g in s.get("ex_groups", []) if g in groups]
            st.multiselect("분류 (비워 두면 전체)", groups, key="ex_groups", placeholder="전체 (눌러서 고르기)")
        else:
            st.caption("분류가 적힌 한자가 없어서 전체에서 나와요.")
    pool = exam_pool()
    lim = exam_limits(pool)
    why = exam_blocked_reasons(lim, pool)
    st.markdown("**유형별 문항 수와 배점**")
    h = st.columns([5, 3, 3])
    h[0].caption("유형")
    h[1].caption("문항 수")
    h[2].caption("문항당 배점")
    for k in EXAM_ORDER:
        off = lim[k] == 0
        r = st.columns([5, 3, 3], vertical_alignment="center")
        r[0].write(EXAM_TYPES[k] + (" 🔒" if off else ""))
        r[1].number_input("문항 수", 0, 100, step=1, key=f"ex_n_{k}", label_visibility="collapsed", disabled=off)
        r[2].number_input("배점", 0.0, 100.0, step=0.5, format="%.1f", key=f"ex_pt_{k}",
                          label_visibility="collapsed", disabled=off)
        if off and k in why:
            st.caption(f"🔒 {EXAM_TYPES[k]}: {why[k]}")
    n_all, tot = exam_total(lim)
    need = sum(0 if lim[k] == 0 else _int(s.get(f"ex_n_{k}"), 0, 0, 100) for k in ("ko", "write", "choice"))
    if need > len(pool):
        st.warning(f"범위 안 한자가 {len(pool)}개라서 한자 문항 {need}개를 다 채울 수 없어요. 모자란 만큼 줄어들어요.")
    m1, m2 = st.columns(2)
    m1.metric("총 문항", f"{n_all}개")
    m2.metric("배점 합계", f"{fmt_pts(tot)}점")
    if n_all and abs(tot - 100.0) > 0.05:
        st.warning(f"배점 합계가 {fmt_pts(tot)}점이에요. 100점 만점이 아니에요.")
        st.button("⚖️ 100점으로 맞추기", on_click=fit_points, width="stretch", key="ex_fit",
                  help="지금 배점의 비율은 그대로 두고 합계만 100점이 되게 조정해요. (소수 첫째 자리)")
        st.caption("소수 첫째 자리로 반올림하기 때문에 정확히 100점이 안 될 수도 있어요. 그땐 배점을 직접 조정해 주세요.")
    if s.get("ex_apply") not in ["반영해요", "반영 안 해요"]:
        s.ex_apply = "반영해요"
    st.radio("시험이 끝나면 결과를 오답 노트·복습 기록에 …", ["반영해요", "반영 안 해요"], key="ex_apply", horizontal=True,
             help="'반영해요'를 고르면 틀린 문제는 오답 노트에 담기고, 맞고 틀림이 복습 주기와 통계에도 들어가요.")
    st.caption("사자성어 문항은 범위와 상관없이 전체 사자성어에서 나와요. (범위가 '오답 노트만'이면 사자성어 오답에서 나와요)")
    st.button("▶️ 시험 시작", type="primary", on_click=start_exam, disabled=(n_all == 0), width="stretch", key="ex_start")
    if s.exams:
        with st.expander("📈 지난 모의고사 기록"):
            st.caption(f"지금까지 {len(s.exams)}회")
            render_exam_history()


def render_exam():
    s = st.session_state
    if s.exam:
        render_exam_run()
    elif s.exam_result:
        render_exam_result()
    else:
        render_exam_setup()


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
            if s.exams:
                last = s.exams[-1]
                st.caption(f"최근 모의고사 {fmt_pts(exam_pct(last))}점 ({last['date']})")
            if p["D"] <= 2:
                st.caption("시험 직전 이틀은 새 한자 없이 복습만 해요.")
            n_today = p["review"] + p["new_today"]
            st.button(f"▶️ 오늘 플랜 시작 (복습 {p['review']} + 새 한자 {p['new_today']})", type="primary",
                      on_click=start_plan, disabled=(n_today == 0), width="stretch", key="plan_start")
            if n_today == 0:
                st.caption("오늘 할 몫을 다 했어요. 🎉 '시험' 탭에서 모의고사로 실력을 확인해 보세요.")
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
        st.caption("별표를 준 한자는 연습·카드의 '별표 한자만' 범위로 풀 수 있고, 모의고사 범위로도 골라요.")


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


def render_set_tools():
    with st.expander("📂 내장 단어장 불러오기"):
        set_picker("word", "study", show=False)
    render_word_edit()



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
