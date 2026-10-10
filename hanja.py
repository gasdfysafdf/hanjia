"""한자 퀴즈 웹앱 (Streamlit)

- 한자 파일(csv / xlsx / txt) 업로드 -> 문제 자동 생성
- 문제 유형: 뜻·음 쓰기(채팅 입력), 한자 쓰기(캔버스 손글씨)
- 오답 노트, 오답만 다시 풀기, 통계, 암기 모드, 진행 상황 저장/불러오기
"""
import glob
import io
import json
import os
import random
import re
import tempfile
import urllib.request
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

CANVAS_ERR = ""
try:
    from streamlit_drawable_canvas import st_canvas

    HAS_CANVAS = True
except Exception as _e:  # 라이브러리가 없어도 뜻·음 문제는 풀 수 있게
    HAS_CANVAS = False
    CANVAS_ERR = f"{type(_e).__name__}: {_e}"

st.set_page_config(page_title="한자 퀴즈", page_icon="📖", layout="wide")

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
    cols = {"hanja": 0, "meaning": 1, "sound": 2}
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

        rows.append((get(cols["hanja"]), get(cols["meaning"]), get(cols["sound"])))
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


def build_words(rows):
    words, seen, dropped = [], set(), 0
    for h, m, s_ in rows:
        h = h.strip()
        if not has_hanja(h) or not s_.strip():
            dropped += 1
            continue
        if h in seen:
            continue
        seen.add(h)
        words.append({"hanja": h, "meaning": m.strip(), "sound": s_.strip()})
    return words, dropped


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


def guide_image(ch, font_path, size=320):
    mask = np.array(render_char(ch, font_path, size)) > 127
    base = np.full((size, size), 255, np.uint8)
    base[mask] = 215  # 연한 회색 (채점 때 획으로 인식되지 않는 밝기)
    return Image.fromarray(base).convert("RGB")


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
def init_state():
    defaults = dict(
        words=[], file_sig=None, quiz=[], pos=0, results=[], feedback=None,
        wrong={}, stats={}, nonce=0, retry_mode=False, hint=False, flash="",
        progress_sig=None, calib=[], set_preset="보통",
    )
    defaults.update(GRADE_DEFAULTS)
    for k, v in defaults.items():
        st.session_state.setdefault(k, list(v) if isinstance(v, list) else v)


def reset_round(quiz, retry=False):
    s = st.session_state
    s.quiz, s.pos, s.results = quiz, 0, []
    s.feedback, s.hint, s.flash = None, False, ""
    s.retry_mode = retry
    s.nonce += 1


def start_quiz():
    s = st.session_state
    source = s.get("set_source", "전체 단어")
    pool = list(s.wrong.values()) if source == "오답 노트만" else list(s.words)
    if not pool:
        s.flash = "문제로 만들 단어가 없어요. 파일을 올리거나 오답 노트를 확인해 주세요."
        return
    kind = s.get("set_type", "섞어서")
    note = ""
    if kind in ("한자 쓰기", "섞어서") and not HAS_CANVAS:
        if kind == "한자 쓰기":
            s.flash = (
                "⚠️ 한자 쓰기 기능(streamlit-drawable-canvas)을 불러오지 못해서 시작할 수 없어요. "
                "왼쪽 아래 '진단 정보'를 확인해 주세요."
            )
            return
        note = "⚠️ 한자 쓰기 기능을 불러오지 못해 뜻·음 문제만 나와요. (왼쪽 '진단 정보' 확인)"
    items = []
    for w in pool:
        can_write = HAS_CANVAS and len(w["hanja"]) == 1
        if kind == "뜻·음 쓰기":
            ty = "read"
        elif kind == "한자 쓰기":
            ty = "write" if can_write else None
        else:
            ty = random.choice(["read", "write"]) if can_write else "read"
        if ty:
            items.append({"w": w, "type": ty})
    if s.get("set_shuffle", True):
        random.shuffle(items)
    items = items[: int(s.get("set_count", 10))]
    if not items:
        s.flash = "선택한 유형으로 만들 수 있는 문제가 없어요. (한자 쓰기는 한 글자 한자만 가능)"
        return
    reset_round(items, retry=(source == "오답 노트만"))
    s.flash = note


def record(w, qtype, ok, ans):
    s = st.session_state
    key = w["hanja"]
    entry = {"w": w, "type": qtype, "correct": ok, "answer": ans, "added_wrong": False}
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
        e["last"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        entry["added_wrong"] = True
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
    s.flash = "퀴즈를 초기화했어요. '시작'을 눌러 새로 시작하세요."


def reset_all():
    """퀴즈 + 오답 노트 + 통계 전부 초기화 (불러온 단어는 유지)"""
    s = st.session_state
    s.wrong, s.stats = {}, {}
    reset_quiz()
    s.confirm_reset = False
    s.flash = "퀴즈, 오답 노트, 통계를 모두 초기화했어요."


# ----------------------------------------------------------------------
# 화면: 문제 풀기
# ----------------------------------------------------------------------
def big(text, size=96):
    st.markdown(
        f"<div style='font-size:{size}px;text-align:center;line-height:1.2'>{text}</div>",
        unsafe_allow_html=True,
    )


def render_read(w):
    s = st.session_state
    with st.chat_message("assistant", avatar="📖"):
        big(w["hanja"])
        st.write("이 한자의 **뜻과 음**을 입력하세요. 예) 하늘 천")
        if s.hint:
            first_mean = split_alts(w["meaning"])[0] if split_alts(w["meaning"]) else ""
            st.caption(f"💡 음은 '{w['sound'][0]}'(으)로 시작하고, 뜻은 {len(first_mean)}글자예요.")
    c1, c2, _ = st.columns([1, 1, 4])
    if c1.button("💡 힌트"):
        s.hint = True
        st.rerun()
    if c2.button("모르겠어요", key="skip_read"):
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


def draw_canvas(bg, key):
    s = st.session_state
    kwargs = dict(
        fill_color="rgba(255,255,255,0)",
        stroke_width=int(s.get("set_stroke", 8)),
        stroke_color="#000000",
        background_color="#FFFFFF",
        height=320,
        width=320,
        drawing_mode="freedraw",
        key=key,
    )
    err = None
    try:
        return st_canvas(background_image=bg, **kwargs)
    except Exception as e:
        err = e
    if bg is not None:  # 가이드(배경 이미지)가 문제일 수 있어 끄고 재시도
        st.warning("이 환경에서는 연한 글씨 가이드가 동작하지 않아 끄고 진행해요.")
        try:
            return st_canvas(background_image=None, **kwargs)
        except Exception as e2:
            err = e2
    st.error(f"캔버스를 표시하지 못했어요: {type(err).__name__}: {err}")
    return None


def render_write(w):
    s = st.session_state
    font = find_font()
    st.markdown(f"### ✍️ **{w['meaning']} {w['sound']}** 에 해당하는 한자를 쓰세요")
    hint = st.checkbox("연한 글씨 가이드 보기 (힌트)", key=f"hint_{s.pos}_{s.nonce}", disabled=font is None)
    bg = guide_image(w["hanja"], font) if (hint and font) else None
    left, _ = st.columns([1, 1])
    with left:
        res = draw_canvas(bg, f"cv_{s.pos}_{s.nonce}_{int(hint)}")
    b1, b2, b3, _ = st.columns([1, 1, 1, 3])
    if b1.button("제출", type="primary"):
        data = res.image_data if res is not None else None
        if data is None or (data[:, :, :3].astype(float).mean(axis=2) < 110).sum() < 30:
            st.warning("먼저 한자를 써 주세요.")
        else:
            submit_write(w, data, font)
    if b2.button("지우기"):
        s.nonce += 1
        st.rerun()
    if b3.button("모르겠어요", key="skip_write"):
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


def render_feedback(q):
    s = st.session_state
    fb, w = s.feedback, q["w"]
    if fb["type"] == "read":
        with st.chat_message("user"):
            st.write(fb["user"])
        with st.chat_message("assistant", avatar="📖"):
            if fb["ok"]:
                st.success("정답! 🎉")
            else:
                st.error("오답이에요.")
            st.markdown(f"정답: **{w['hanja']} = {w['meaning']} {w['sound']}**")
            if fb.get("detail"):
                st.caption(fb["detail"])
    else:
        big(w["hanja"], 110)
        st.markdown(f"<div style='text-align:center'>{w['meaning']} {w['sound']}</div>", unsafe_allow_html=True)
        if fb["pending"]:
            st.info("정답 한자와 내 글씨를 비교해서 직접 채점해 주세요. (서버에 한자 폰트가 없어 자동 채점이 꺼져 있어요)")
            g1, g2, _ = st.columns([1, 1, 4])
            if g1.button("⭕ 맞았어요"):
                self_grade(w, True)
            if g2.button("❌ 틀렸어요"):
                self_grade(w, False)
            return
        if fb["ok"]:
            st.success("정답! 🎉" + (f"  (유사도 {fb['score']:.2f})" if fb["score"] is not None else ""))
        else:
            st.error("오답이에요." + (f"  (유사도 {fb['score']:.2f})" if fb["score"] is not None else ""))
        if fb.get("p") is not None:
            line = (f"내 글씨가 정답과 맞는 정도 {fb['p']:.0%} · 정답 획을 채운 정도 {fb['r']:.0%}"
                    f" · 가장 크게 어긋난 획 {fb['gap']:.0f}%")
            if fb.get("other"):
                line += f" · 가장 비슷한 다른 글자 '{fb['other'][0]}' {fb['other'][1]:.2f}"
            st.caption(line)
        if fb.get("overlay") is not None:
            st.image(fb["overlay"], caption="빨강: 정답 글자 / 파랑: 내 글씨 / 보라: 겹친 부분", width=256)
        if fb.get("detail"):
            st.caption(fb["detail"])
    if not fb["ok"] and not fb.get("pending"):
        st.button("🙋 내가 맞게 썼어요 (정정)", on_click=override_correct)
    st.button("다음 문제 ➡️", type="primary", on_click=next_question)


def render_summary():
    s = st.session_state
    res = s.results
    n = len(res)
    c = sum(r["correct"] for r in res)
    st.subheader("🎉 결과")
    m1, m2, m3 = st.columns(3)
    m1.metric("점수", f"{c} / {n}")
    m2.metric("정답률", f"{(c / n * 100 if n else 0):.0f}%")
    m3.metric("오답 노트 누적", len(s.wrong))
    if n:
        df = pd.DataFrame(
            [
                {
                    "결과": "⭕" if r["correct"] else "❌",
                    "유형": "뜻·음" if r["type"] == "read" else "한자 쓰기",
                    "한자": r["w"]["hanja"],
                    "뜻": r["w"]["meaning"],
                    "음": r["w"]["sound"],
                    "내 답": r["answer"],
                }
                for r in res
            ]
        )
        st.dataframe(df, width="stretch", hide_index=True)
    wrong_items = [{"w": r["w"], "type": r["type"]} for r in res if not r["correct"]]
    b1, b2, _ = st.columns([1, 1.4, 3])
    b1.button("🔁 새로 시작", on_click=start_quiz)
    if wrong_items:
        b2.button("❌ 이번에 틀린 것만 다시", on_click=reset_round, args=(wrong_items, False))


# ----------------------------------------------------------------------
# 화면: 오답 노트 / 통계 / 단어장
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
    c1, c2, c3, c4 = st.columns(4)
    if c1.button("선택 삭제") and sel:
        for k in sel:
            s.wrong.pop(k, None)
        st.rerun()
    if c2.button("전체 비우기"):
        s.wrong.clear()
        st.rerun()
    c3.download_button("CSV로 저장", df.to_csv(index=False).encode("utf-8-sig"), "오답노트.csv", "text/csv")

    def go_wrong():
        s.set_source = "오답 노트만"
        start_quiz()

    c4.button("📝 오답으로 바로 풀기", on_click=go_wrong)
    st.caption("오답 노트 모드에서 맞히면 노트에서 자동으로 빠져요. 풀이는 '문제 풀기' 탭에서 진행돼요.")


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
    tr, tw = df["맞음"].sum(), df["틀림"].sum()
    a, b, c = st.columns(3)
    a.metric("푼 문제", int(tr + tw))
    b.metric("전체 정답률", f"{tr / (tr + tw) * 100:.0f}%")
    c.metric("약한 한자(정답률 50% 미만)", int((df["정답률(%)"] < 50).sum()))
    st.dataframe(df, width="stretch", hide_index=True)
    st.bar_chart(df.set_index("한자")[["맞음", "틀림"]])


def render_study():
    s = st.session_state
    if not s.words:
        st.info("파일을 업로드하면 단어장이 나와요.")
        return
    hide = st.toggle("암기 모드 (뜻·음 가리기)", value=False)
    df = pd.DataFrame(s.words).rename(columns={"hanja": "한자", "meaning": "뜻", "sound": "음"})
    if hide:
        df["뜻"], df["음"] = "???", "?"
    st.dataframe(df, width="stretch", hide_index=True, height=500)


# ----------------------------------------------------------------------
# 사이드바 + 메인
# ----------------------------------------------------------------------
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


SAMPLE_CSV = "한자,뜻,음\n天,하늘,천\n地,땅,지\n人,사람,인\n山,메,산\n水,물,수\n火,불,화\n日,날,일\n月,달,월\n"


def sidebar():
    s = st.session_state
    st.sidebar.title("📖 한자 퀴즈")
    f = st.sidebar.file_uploader("한자 파일 업로드 (txt / csv / xlsx)")  # 형식 제한 없음: 태블릿 호환
    st.sidebar.download_button("샘플 CSV 받기", SAMPLE_CSV.encode("utf-8-sig"), "sample_hanja.csv", "text/csv")
    if f is not None:
        sig = (f.name, f.size)
        if sig != s.file_sig:
            try:
                words, dropped = parse_file(f)
                s.words, s.file_sig = words, sig
                s.quiz = []
                s.flash = f"{len(words)}개 한자를 불러왔어요." + (
                    f" (뜻·음이 없어서 제외한 줄 {dropped}개)" if dropped else ""
                )
            except Exception as e:
                s.flash = f"파일을 읽지 못했어요: {e}"
    with st.sidebar.expander("✍️ 텍스트 직접 붙여넣기 (파일이 안 될 때)"):
        pasted = st.text_area("한 줄에 '한자 뜻 음' (예: 天 하늘 천)", height=150, key="paste_text")
        if st.button("붙여넣은 내용 적용"):
            words, dropped = build_words(parse_text(pasted))
            if words:
                s.words, s.quiz = words, []
                s.file_sig = (f.name, f.size) if f is not None else None  # 업로드 파일이 덮어쓰지 않게
                s.flash = f"{len(words)}개 한자를 불러왔어요."
            else:
                s.flash = "읽을 수 있는 줄이 없어요. 한 줄에 '한자 뜻 음' 모양인지 확인해 주세요."
    st.sidebar.caption(f"불러온 한자: {len(s.words)}개")

    st.sidebar.divider()
    st.sidebar.radio("문제 유형", ["섞어서", "뜻·음 쓰기", "한자 쓰기"], key="set_type")
    st.sidebar.radio("문제 범위", ["전체 단어", "오답 노트만"], key="set_source")
    st.sidebar.number_input("문제 수", 1, 500, 10, key="set_count")
    st.sidebar.checkbox("순서 섞기", value=True, key="set_shuffle")
    grading_sidebar()
    st.sidebar.slider("펜 굵기", 3, 20, 8, key="set_stroke")
    st.sidebar.button("▶️ 시작", type="primary", on_click=start_quiz, width="stretch")

    if not find_font():
        st.sidebar.warning("한자 폰트를 찾지 못하고 자동으로 내려받지도 못해서 손글씨는 자가 채점으로 진행돼요. (packages.txt에 fonts-noto-cjk가 있는지 확인)")
    if not HAS_CANVAS:
        st.sidebar.error(
            "⚠️ 한자 쓰기 기능(streamlit-drawable-canvas)을 불러오지 못했어요. "
            "저장소의 requirements.txt(소문자)에 streamlit-drawable-canvas==0.9.3 이 있는지 확인하고 앱을 Reboot 해 주세요."
        )
        st.sidebar.code(CANVAS_ERR)
    with st.sidebar.expander("진단 정보"):
        st.write(f"Streamlit 버전: {st.__version__}")
        st.write(f"캔버스 라이브러리: {'정상' if HAS_CANVAS else '불러오기 실패'}")
        st.write(f"한자 폰트: {find_font() or '없음'} ({_FONT['source'] or '-'})")

    st.sidebar.divider()
    st.sidebar.button("🔄 퀴즈 초기화", on_click=reset_quiz, width="stretch")
    st.sidebar.checkbox("전체 초기화 확인 (오답 노트·통계 삭제)", key="confirm_reset")
    st.sidebar.button(
        "🗑️ 전체 초기화",
        on_click=reset_all,
        disabled=not s.get("confirm_reset", False),
        width="stretch",
    )

    st.sidebar.divider()
    st.sidebar.caption("진행 상황 저장 / 불러오기 (오답 노트 + 통계)")
    payload = json.dumps({"wrong": s.wrong, "stats": s.stats}, ensure_ascii=False, indent=2)
    st.sidebar.download_button("💾 저장", payload.encode("utf-8"), "hanja_progress.json", "application/json")
    pf = st.sidebar.file_uploader("불러오기", type=["json"], key="progress_file")
    if pf is not None and (pf.name, pf.size) != s.progress_sig:
        try:
            data = json.loads(pf.getvalue().decode("utf-8"))
            s.wrong.update(data.get("wrong", {}))
            for k, v in data.get("stats", {}).items():
                s.stats[k] = v
            s.progress_sig = (pf.name, pf.size)
            s.flash = "진행 상황을 불러왔어요."
        except Exception:
            s.flash = "진행 파일을 읽지 못했어요."


def main():
    init_state()
    sidebar()
    s = st.session_state
    if s.flash:
        st.info(s.flash)
    if not HAS_CANVAS:
        st.warning("한자 쓰기 기능을 불러오지 못해 지금은 뜻·음 문제만 풀 수 있어요. 왼쪽 '진단 정보'를 확인해 주세요.")
    t1, t2, t3, t4 = st.tabs(["✏️ 문제 풀기", "📕 오답 노트", "📊 통계", "📚 단어장"])
    with t1:
        if not s.quiz:
            st.write("왼쪽에서 파일을 올리고 **시작**을 눌러 주세요.")
        elif s.pos >= len(s.quiz):
            render_summary()
        else:
            q = s.quiz[s.pos]
            st.progress(s.pos / len(s.quiz), text=f"{s.pos + 1} / {len(s.quiz)}")
            done = len(s.results)
            st.caption(f"현재 점수 {sum(r['correct'] for r in s.results)} / {done}")
            if s.feedback:
                render_feedback(q)
            elif q["type"] == "read":
                render_read(q["w"])
            else:
                render_write(q["w"])
    with t2:
        render_wrong()
    with t3:
        render_stats()
    with t4:
        render_study()


main()
