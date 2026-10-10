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
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFilter, ImageFont

try:
    from streamlit_drawable_canvas import st_canvas

    HAS_CANVAS = True
except Exception:  # 라이브러리가 없어도 뜻·음 문제는 풀 수 있게
    HAS_CANVAS = False

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


@st.cache_resource
def find_font():
    for pat in FONT_PATTERNS:
        if not pat:
            continue
        for p in glob.glob(pat):
            try:
                load_font(p, 40)
                return p
            except Exception:
                continue
    return None


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
# 손글씨 비교 (이미지 유사도)
# ----------------------------------------------------------------------
def render_char(ch, font_path, size=256):
    font = load_font(font_path, int(size * 0.8))
    img = Image.new("L", (size, size), 0)
    ImageDraw.Draw(img).text((size / 2, size / 2), ch, font=font, fill=255, anchor="mm")
    return img


def guide_image(ch, font_path, size=320):
    mask = np.array(render_char(ch, font_path, size)) > 127
    base = np.full((size, size), 255, np.uint8)
    base[mask] = 215  # 연한 회색 (채점 때 획으로 인식되지 않는 밝기)
    return Image.fromarray(base).convert("RGB")


def normalize(img, out=64, pad=4):
    arr = np.array(img) > 127
    ys, xs = np.where(arr)
    if len(ys) == 0:
        return None
    crop = img.crop((xs.min(), ys.min(), xs.max() + 1, ys.max() + 1))
    w, h = crop.size
    side = max(w, h)
    sq = Image.new("L", (side, side), 0)
    sq.paste(crop, ((side - w) // 2, (side - h) // 2))
    inner = out - 2 * pad
    sq = sq.resize((inner, inner), Image.BILINEAR)
    canvas = Image.new("L", (out, out), 0)
    canvas.paste(sq, (pad, pad))
    return np.array(canvas) > 100


def dilate(mask, size=7):
    img = Image.fromarray((mask * 255).astype("uint8"))
    return np.array(img.filter(ImageFilter.MaxFilter(size))) > 127


def similarity(u, t):
    ud, td = dilate(u), dilate(t)
    p = (u & td).sum() / max(u.sum(), 1)
    r = (t & ud).sum() / max(t.sum(), 1)
    return 0.0 if p + r == 0 else float(2 * p * r / (p + r))


def overlay(u, t, scale=4):
    h, w = u.shape
    rgb = np.full((h, w, 3), 255, np.uint8)
    rgb[t] = (255, 190, 190)
    rgb[u] = (30, 60, 220)
    rgb[u & t] = (120, 40, 160)
    return Image.fromarray(rgb).resize((w * scale, h * scale), Image.NEAREST)


def compare_writing(image_data, ch, font_path):
    gray = image_data[:, :, :3].astype(float).mean(axis=2)
    user_img = Image.fromarray(((gray < 110) * 255).astype("uint8"))
    u = normalize(user_img)
    t = normalize(render_char(ch, font_path))
    if u is None or t is None:
        return None
    return similarity(u, t), u, t


# ----------------------------------------------------------------------
# 상태 관리
# ----------------------------------------------------------------------
def init_state():
    defaults = dict(
        words=[], file_sig=None, quiz=[], pos=0, results=[], feedback=None,
        wrong={}, stats={}, nonce=0, retry_mode=False, hint=False, flash="",
        progress_sig=None,
    )
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


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


def render_write(w):
    s = st.session_state
    font = find_font()
    st.markdown(f"### ✍️ **{w['meaning']} {w['sound']}** 에 해당하는 한자를 쓰세요")
    hint = st.checkbox("연한 글씨 가이드 보기 (힌트)", key=f"hint_{s.pos}_{s.nonce}", disabled=font is None)
    bg = guide_image(w["hanja"], font) if (hint and font) else None
    left, _ = st.columns([1, 1])
    with left:
        res = st_canvas(
            fill_color="rgba(255,255,255,0)",
            stroke_width=int(s.get("set_stroke", 8)),
            stroke_color="#000000",
            background_color="#FFFFFF",
            background_image=bg,
            height=320,
            width=320,
            drawing_mode="freedraw",
            key=f"cv_{s.pos}_{s.nonce}_{int(hint)}",
        )
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
        r = compare_writing(data, w["hanja"], font)
        if r is None:
            st.warning("글씨를 인식하지 못했어요. 조금 더 크게 써 주세요.")
            return
        score, u, t = r
        ok = score >= s.get("set_thr", 0.6)
        record(w, "write", ok, f"유사도 {score:.2f}")
        s.feedback = {"type": "write", "ok": ok, "pending": False, "score": score, "overlay": overlay(u, t)}
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
        st.dataframe(df, use_container_width=True, hide_index=True)
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
    st.dataframe(df, use_container_width=True, hide_index=True)
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
    st.dataframe(df, use_container_width=True, hide_index=True)
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
    st.dataframe(df, use_container_width=True, hide_index=True, height=500)


# ----------------------------------------------------------------------
# 사이드바 + 메인
# ----------------------------------------------------------------------
SAMPLE_CSV = "한자,뜻,음\n天,하늘,천\n地,땅,지\n人,사람,인\n山,메,산\n水,물,수\n火,불,화\n日,날,일\n月,달,월\n"


def sidebar():
    s = st.session_state
    st.sidebar.title("📖 한자 퀴즈")
    f = st.sidebar.file_uploader("한자 파일 업로드 (csv / xlsx / txt)", type=["csv", "xlsx", "xls", "txt"])
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
    st.sidebar.caption(f"불러온 한자: {len(s.words)}개")

    st.sidebar.divider()
    st.sidebar.radio("문제 유형", ["섞어서", "뜻·음 쓰기", "한자 쓰기"], key="set_type")
    st.sidebar.radio("문제 범위", ["전체 단어", "오답 노트만"], key="set_source")
    st.sidebar.number_input("문제 수", 1, 500, 10, key="set_count")
    st.sidebar.checkbox("순서 섞기", value=True, key="set_shuffle")
    st.sidebar.slider("손글씨 채점 기준 (높을수록 엄격)", 0.3, 0.9, 0.6, 0.05, key="set_thr")
    st.sidebar.slider("펜 굵기", 3, 20, 8, key="set_stroke")
    st.sidebar.button("▶️ 시작", type="primary", on_click=start_quiz, use_container_width=True)

    if not find_font():
        st.sidebar.warning("한자 폰트를 못 찾아서 손글씨는 자가 채점으로 진행돼요. (README 참고)")
    if not HAS_CANVAS:
        st.sidebar.warning("streamlit-drawable-canvas가 없어 '한자 쓰기' 문제는 꺼져 있어요.")

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
