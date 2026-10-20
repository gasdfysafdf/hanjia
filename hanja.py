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
import os
import random
import re
import tempfile
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


FONT_URL = "https://raw.githubusercontent.com/notofonts/noto-cjk/main/Sans/SubsetOTF/KR/NotoSansKR-Regular.otf"
_FONT = {"path": None, "source": ""}


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
        if not (os.path.exists(path) and os.path.getsize(path) > 1_000_000):
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with urllib.request.urlopen(FONT_URL, timeout=60) as r, open(path + ".part", "wb") as f:
                f.write(r.read())
            os.replace(path + ".part", path)
        load_font(path, 40)
        return path
    except Exception:
        return None


def find_font():
    """성공한 결과만 저장해서, 실패하면 다음 실행 때 다시 시도"""
    if _FONT["path"]:
        return _FONT["path"]
    p = _system_font()
    if p:
        _FONT.update(path=p, source="서버에 설치된 폰트")
        return p
    p = _download_font()
    if p:
        _FONT.update(path=p, source="자동 내려받음")
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


def parse_text(text):
    rows = []
    for line in text.splitlines():
        toks = [t for t in re.split(r"[,\t;|\s]+", line.strip()) if t]
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
    if not any(has_hanja(x) for x in first):  # 첫 줄이 헤더
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
                io.StringIO(decode_bytes(data)), header=None, dtype=str, keep_default_na=False
            )
            rows = parse_table(df)
        except Exception:
            rows = parse_text(decode_bytes(data))
    else:
        rows = parse_text(decode_bytes(data))

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
    set_goal=20, set_stroke=8,
    card_src="전체 단어", card_dir="한자 → 뜻·음", card_shuffle=True, card_count=20,
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


def start_quiz():
    s = st.session_state
    source = s.get("set_source", "전체 단어")
    kind = s.get("set_type", "섞어서")
    groups = s.get("set_groups", [])
    if source == "오답 노트만":
        pool = list(s.wrong.values())
    else:
        pool = [w for w in s.words if not groups or w.get("group") in groups]
        if source == "오늘 복습할 것":
            pool = due_pool(pool)
    if not pool:
        if source == "오늘 복습할 것":
            s.flash = "오늘 복습할 한자가 없어요. 내일 다시 만나요! 🎉"
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
    count = int(s.get("set_count", 10))
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
    s.cards = None
    reset_quiz()
    s.confirm_reset = False
    s.flash = "퀴즈, 오답 노트, 통계, 복습 기록을 모두 초기화했어요."


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
    "set_goal", "set_stroke", "set_type", "set_source", "set_count", "set_shuffle", "set_preset",
    "card_src", "card_dir", "card_shuffle", "card_count",
] + list(GRADE_DEFAULTS)
ALLOWED = {
    "set_type": ["섞어서", "뜻·음 쓰기", "한자 쓰기", "객관식"],
    "set_source": ["전체 단어", "오늘 복습할 것", "오답 노트만"],
    "card_src": ["전체 단어", "오늘 복습할 것", "오답 노트만"],
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
    out = {"wrong": {}, "stats": {}, "srs": {}, "history": {}}
    if not isinstance(d, dict):
        return out
    for k, v in (d.get("wrong") or {}).items():
        if isinstance(v, dict) and has_hanja(str(k)):
            out["wrong"][str(k)] = {
                "hanja": str(k), "meaning": str(v.get("meaning", "")), "sound": str(v.get("sound", "")),
                "wrong_count": _int(v.get("wrong_count"), 1, 1), "last": str(v.get("last", "")),
            }
    for k, v in (d.get("stats") or {}).items():
        if isinstance(v, dict) and has_hanja(str(k)):
            out["stats"][str(k)] = {
                "meaning": str(v.get("meaning", "")), "sound": str(v.get("sound", "")),
                "right": _int(v.get("right")), "wrong": _int(v.get("wrong")),
            }
    for k, v in (d.get("srs") or {}).items():
        due = _iso(v.get("due")) if isinstance(v, dict) else None
        if due and has_hanja(str(k)):
            out["srs"][str(k)] = {"box": _int(v.get("box"), 0, 0, len(INTERVALS) - 1), "due": due,
                                  "seen": _int(v.get("seen"))}
    for day, v in (d.get("history") or {}).items():
        if isinstance(v, dict) and _iso(day):
            out["history"][_iso(day)] = {"right": _int(v.get("right")), "wrong": _int(v.get("wrong"))}
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


def build_payload():
    s = st.session_state
    data = {"v": 1, "words": s.words, "wrong": s.wrong, "stats": s.stats, "srs": s.srs,
            "history": s.history, "settings": {k: s[k] for k in PERSIST_KEYS if k in s}}
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
    merge_progress(sanitize_progress(data))
    defaults = dict(GRADE_DEFAULTS)
    defaults.update(SETTING_DEFAULTS)
    defaults["set_preset"] = "보통"
    for k, v in (data.get("settings") or {}).items():
        if k in PERSIST_KEYS and k in defaults and type(v) is type(defaults[k]):
            if k in ALLOWED and v not in ALLOWED[k]:
                continue
            s[k] = v
    return bool(s.words or s.wrong or s.stats or s.srs or s.history)


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
    if s.restored and (s.words or s.wrong or s.stats or s.srs or s.history):
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
    else:
        pool = list(s.words)
    if not pool:
        s.flash = "오늘 복습할 한자가 없어요. 🎉" if src == "오늘 복습할 것" else "카드로 만들 한자가 없어요."
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

/* 결과 요약 */
.score-box{text-align:center; padding:1.2rem; border:2px solid var(--indigo); border-radius:6px; background:#fff; margin-bottom:1rem;}
.score-box .big{font:700 3.2rem/1.1 var(--serif); color:var(--indigo);}
.score-box .lbl{color:var(--soft);}

/* 어두운 화면 설정(다크 모드)에서도 글씨가 보이게: 흰 배경 위의 글씨는 색을 직접 지정 */
.stRadio [role="radiogroup"] label, .stRadio [role="radiogroup"] label p{color:var(--ink);}
.st-key-choices_ko button, .st-key-choices_hj button, .st-key-choices_ko button p, .st-key-choices_hj button p{color:var(--ink);}

/* 자동 저장용 보이지 않는 조각 */
.st-key-store_box{height:0; margin:0 0 -1rem; padding:0; overflow:hidden;}

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
    if review:
        chips.append(f"<span class='chip hot'>🔁 복습할 차례 {review}개</span>")
    if new:
        chips.append(f"<span class='chip'>🌱 아직 안 푼 {new}개</span>")
    if s.wrong:
        chips.append(f"<span class='chip'>📕 오답 노트 {len(s.wrong)}개</span>")
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
    load_panel()
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
    st.subheader("오늘은 무엇을 연습할까요?")
    new, review = due_counts(s.words)
    if review + new:
        st.button(f"⚡ 오늘 복습 바로 시작 (복습 {review} · 새 한자 {new})", on_click=quick_start, width="stretch")
    st.radio("문제 유형", ["섞어서", "뜻·음 쓰기", "한자 쓰기", "객관식"], key="set_type", horizontal=True,
             help="뜻·음 쓰기: 키보드로 입력 / 한자 쓰기: 손글씨 / 객관식: 4개 중 고르기")
    st.radio("문제 범위", ["전체 단어", "오늘 복습할 것", "오답 노트만"], key="set_source", horizontal=True,
             help=f"'오늘 복습할 것'은 복습 날짜가 된 한자({review}개)와 아직 안 푼 한자({new}개)를 약한 것부터 골라줘요.")
    groups = sorted({w["group"] for w in s.words if w.get("group")})
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
CANVAS_PX = 360


def flatten(data):
    """캔버스 이미지(RGBA)를 흰 바탕에 합성해요. 배경이 투명해도 빈 곳이 '검은 획'으로 읽히지 않게 하는 안전장치"""
    a = data[:, :, 3:4].astype(float) / 255.0
    rgb = data[:, :, :3].astype(float) * a + 255.0 * (1.0 - a)
    out = np.dstack([rgb, np.full(data.shape[:2], 255.0)])
    return out.astype(np.uint8)


def canvas_backdrop(img):
    """캔버스 뒤쪽 칸(iframe)에 십자 점선·힌트 글씨를 CSS 배경으로 깔아요. 캔버스 자체는 투명해서 채점에는 영향이 없어요"""
    buf = io.BytesIO()
    img.save(buf, "PNG")
    uri = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
    html(
        "<style>.st-key-canvas_box{width:412px; max-width:100%; background:url(" + uri + ") 0 0 / "
        f"{CANVAS_PX}px {CANVAS_PX}px no-repeat;}}"
        ".st-key-canvas_box iframe{mix-blend-mode:multiply;}</style>"
    )


def draw_canvas(key):
    s = st.session_state
    kwargs = dict(
        fill_color="rgba(255,255,255,0)",
        stroke_width=int(s.get("set_stroke", 8)),
        stroke_color="#000000",
        background_color="",
        height=CANVAS_PX,
        width=CANVAS_PX,
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
    left, right = st.columns([1, 1])
    with left:
        meaning_card(w)
        html("<div class='ask'>에 해당하는 한자를 오른쪽 칸에 쓰세요</div>")
        hint = st.checkbox("연한 글씨 가이드 보기 (힌트)", key=f"hint_{s.pos}_{s.nonce}", disabled=font is None)
    # 십자 점선은 항상 깔고, 힌트를 켜면 정답 글자가 연한 회색으로 같이 나와요
    canvas_backdrop(guide_image(w["hanja"], font, CANVAS_PX) if (hint and font) else grid_image(CANVAS_PX))
    with right:
        with st.container(key="canvas_box"):
            res = draw_canvas(f"cv_{s.pos}_{s.nonce}_{int(hint)}")
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
        st.radio("카드 범위", ["전체 단어", "오늘 복습할 것", "오답 노트만"], key="card_src", horizontal=True)
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


def render_study():
    s = st.session_state
    q = st.text_input("검색 (한자·뜻·음)", key="study_q", placeholder="예: 하늘, 천, 天")
    c1, c2 = st.columns(2)
    hide = c1.toggle("암기 모드 (뜻·음 가리기, 눌러서 확인)", value=False)
    view = c2.radio("보기", ["카드", "표"], horizontal=True, label_visibility="collapsed")
    words = s.words
    if q.strip():
        qq = q.strip().lower()
        words = [w for w in words if qq in (w["hanja"] + w["meaning"] + w["sound"] + w.get("group", "")).lower()]
    st.caption(f"{len(words)}개")
    if view == "표":
        df = pd.DataFrame(words).rename(columns={"hanja": "한자", "meaning": "뜻", "sound": "음", "group": "분류"})
        if hide and len(df):
            df["뜻"], df["음"] = "???", "?"
        st.dataframe(df, width="stretch", hide_index=True, height=480)
    else:
        cards = []
        for w in words:
            mean = f"<b>{esc(w['meaning'])}</b> {esc(w['sound'])}"
            body = f"<details><summary>보기</summary>{mean}</details>" if hide else mean
            g = f"<div class='g'>{esc(w['group'])}</div>" if w.get("group") else ""
            cards.append(f"<div class='wcard'><div class='h'>{esc(w['hanja'])}</div><div class='m'>{body}</div>{g}</div>")
        html("<div class='wgrid'>" + "".join(cards) + "</div>")
    st.divider()
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
    d, y = page(True)
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
    grading_sidebar()

    with st.sidebar.expander("💾 진행 상황 저장 / 불러오기"):
        st.caption("평소엔 이 기기에 자동 저장돼요. 다른 기기로 옮기거나 백업할 때만 쓰세요. (오답 노트·통계·복습 기록·날짜별 기록)")
        payload = json.dumps(
            {"wrong": s.wrong, "stats": s.stats, "srs": s.srs, "history": s.history}, ensure_ascii=False, indent=2
        )
        st.download_button("💾 저장", payload.encode("utf-8"), "hanja_progress.json", "application/json", width="stretch")
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
        st.divider()
        st.button("🔄 퀴즈 초기화", on_click=reset_quiz, width="stretch")
        st.checkbox("전체 초기화 확인 (오답 노트·통계·복습 기록 삭제)", key="confirm_reset")
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
    t1, t2, t3, t4, t5 = st.tabs(["✏️ 연습", "🃏 카드", "📕 오답 노트", "📊 기록", "📚 단어장"])
    with t1:
        render_practice()
    with t2:
        render_cards()
    with t3:
        render_wrong()
    with t4:
        render_stats()
    with t5:
        render_study()


def main():
    init_state()
    inject_css()  # 먼저 깔아서, 저장용 보이지 않는 조각이 잠깐이라도 보이지 않게
    sync_browser_store()
    try:
        run_app()
    except Exception:  # 예상 못 한 오류가 나도 하얀 화면 대신 복구 버튼을 보여줘요
        st.error("앗, 화면을 그리다가 문제가 생겼어요. 아래 버튼으로 정리하면 대부분 해결돼요. (단어장과 기록은 그대로예요)")
        st.button("🔄 화면 정리하고 다시 시작", on_click=repair_state, type="primary")
        with st.expander("자세한 오류 내용 (문제가 계속되면 이걸 알려 주세요)"):
            st.code(traceback.format_exc())


main()
