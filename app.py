"""
Biopharmaceutics Live Quiz — Streamlit edition
================================================
A Kahoot-style timed quiz for classroom use.

HOW STUDENTS/HOST CONNECT
--------------------------
This one app serves two roles based on a URL parameter:
  - Host (classroom screen):  https://<your-app>.streamlit.app/?role=host
  - Student (on phones):      https://<your-app>.streamlit.app/?role=student

Share the STUDENT link with your class (write it on the board, paste it in
your class group chat, etc). Open the HOST link yourself on the classroom
computer/projector.

EDITING QUESTIONS
------------------
Edit the QUESTIONS list below. Each question needs:
  q             - the question text
  options       - list of exactly 4 answer choices
  correct_index - which option (0-3) is correct
  time          - seconds allowed for that question

No other code needs to change when you edit questions.
"""

import time
import uuid
import streamlit as st
from streamlit_autorefresh import st_autorefresh

# ============================================================
# QUESTION BANK — edit freely
# ============================================================
QUESTIONS = [
    {"q": "Which process best describes the release of drug from its dosage form, making it available for absorption?",
     "options": ["Disintegration", "Dissolution", "Liberation", "Elimination"], "correct_index": 2, "time": 30},
    {"q": "Bioavailability (F) of a drug given by IV bolus is assumed to be:",
     "options": ["0%", "50%", "75%", "100%"], "correct_index": 3, "time": 30},
    {"q": "Which factor does NOT typically affect the rate of drug dissolution?",
     "options": ["Particle size", "Drug pKa/solubility", "Patient's blood type", "Agitation/hydrodynamics"], "correct_index": 2, "time": 30},
    {"q": "First-pass metabolism primarily reduces the bioavailability of drugs administered:",
     "options": ["Intravenously", "Orally", "Sublingually", "Transdermally (in most cases)"], "correct_index": 1, "time": 25},
    {"q": "According to the Noyes-Whitney equation, dissolution rate is directly proportional to:",
     "options": ["Drug's melting point", "Surface area of the particle", "Patient's body weight", "Gastric emptying time"], "correct_index": 1, "time": 30},
    {"q": "A drug with high permeability and low solubility falls into which BCS class?",
     "options": ["Class I", "Class II", "Class III", "Class IV"], "correct_index": 1, "time": 25},
    {"q": "Which parameter describes the fraction of an administered dose that reaches systemic circulation unchanged?",
     "options": ["Clearance", "Volume of distribution", "Bioavailability", "Half-life"], "correct_index": 2, "time": 25},
    {"q": "Enteric coating on a tablet is primarily designed to:",
     "options": ["Speed up gastric dissolution", "Protect the drug from stomach acid or protect stomach from drug", "Improve taste only", "Increase tablet hardness"], "correct_index": 1, "time": 25},
    {"q": "In bioequivalence studies, two products are generally considered bioequivalent if the 90% CI of the AUC and Cmax ratio falls within:",
     "options": ["50-150%", "70-130%", "80-125%", "90-110%"], "correct_index": 2, "time": 30},
    {"q": "Which route of administration avoids first-pass hepatic metabolism entirely?",
     "options": ["Oral", "Rectal (upper)", "Intravenous", "Buccal (partially, but IV is complete)"], "correct_index": 2, "time": 25},
]
REVEAL_SECONDS = 6  # how long the answer reveal / round leaderboard shows before auto-advancing


# ============================================================
# SHARED STATE (one instance shared by every visitor to this app)
# ============================================================
@st.cache_resource
def get_state():
    return {
        "status": "lobby",          # lobby | active | reveal | finished
        "current_q": -1,
        "question_started_at": 0.0,
        "reveal_started_at": 0.0,
        "roster": {},                # pid -> name
        "answers": {},                # q_index -> {pid: {"choice", "correct", "elapsed", "score"}}
    }


def compute_score(correct, elapsed, time_limit):
    if not correct:
        return 0
    raw = 100 - (elapsed / time_limit) * 90
    return max(10, round(raw))


def leaderboard(state):
    totals = {pid: 0 for pid in state["roster"]}
    for qdict in state["answers"].values():
        for pid, a in qdict.items():
            totals[pid] = totals.get(pid, 0) + a["score"]
    rows = [{"name": state["roster"].get(pid, "?"), "score": s} for pid, s in totals.items()]
    rows.sort(key=lambda r: r["score"], reverse=True)
    return rows


def reset_quiz(state):
    state["status"] = "lobby"
    state["current_q"] = -1
    state["question_started_at"] = 0.0
    state["reveal_started_at"] = 0.0
    state["roster"] = {}
    state["answers"] = {}


# ============================================================
# STYLING
# ============================================================
st.set_page_config(page_title="Biopharmaceutics Live Quiz", page_icon="💊", layout="centered")
st.markdown("""
<style>
  html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
  .quiz-card { background:#fff; border:1px solid #D6E2DC; border-radius:18px; padding:32px; }
  .eyebrow { font-family:monospace; font-size:12px; letter-spacing:.14em; text-transform:uppercase; color:#A66E1E; font-weight:600; }
  .big-code { font-family:monospace; font-size:44px; font-weight:700; color:#1F7A6C; text-align:center; margin:10px 0; }
  .qtext { font-size:22px; font-weight:600; text-align:center; margin:18px 0; }
  .board-row { display:flex; justify-content:space-between; padding:10px 14px; border-bottom:1px solid #D6E2DC; }
  .rank { font-family:monospace; font-weight:700; color:#A66E1E; }
  .score { font-family:monospace; font-weight:700; color:#1F7A6C; }
</style>
""", unsafe_allow_html=True)

state = get_state()
role = st.query_params.get("role", "host")


# ============================================================
# HOST VIEW
# ============================================================
def render_host():
    st_autorefresh(interval=1000, key="host_autorefresh")
    st.markdown('<div class="eyebrow">HOST · BIOPHARMACEUTICS LIVE QUIZ</div>', unsafe_allow_html=True)

    if state["status"] == "lobby":
        st.title("Waiting room")
        st.write("Share the **student link** with your class, then start once everyone's in.")
        st.code(f"{_base_url()}/?role=student", language=None)
        names = list(state["roster"].values())
        st.write(f"**{len(names)} student(s) joined**")
        if names:
            st.write(", ".join(names))
        else:
            st.caption("No one yet — waiting...")
        if st.button("Start quiz", disabled=len(names) == 0, type="primary"):
            state["status"] = "active"
            state["current_q"] = 0
            state["question_started_at"] = time.time()
            st.rerun()
        if st.button("Reset session"):
            reset_quiz(state)
            st.rerun()

    elif state["status"] == "active":
        q = QUESTIONS[state["current_q"]]
        elapsed = time.time() - state["question_started_at"]
        remaining = max(0, q["time"] - elapsed)
        st.subheader(f"Question {state['current_q']+1} of {len(QUESTIONS)}")
        st.progress(min(1.0, remaining / q["time"]))
        st.markdown(f"### ⏱ {int(remaining)+1}s")
        st.markdown(f'<div class="qtext">{q["q"]}</div>', unsafe_allow_html=True)
        for i, opt in enumerate(q["options"]):
            st.write(f"{chr(65+i)}. {opt}")
        answered = len(state["answers"].get(state["current_q"], {}))
        st.caption(f"{answered} of {len(state['roster'])} students have answered")
        if remaining <= 0:
            state["status"] = "reveal"
            state["reveal_started_at"] = time.time()
            st.rerun()

    elif state["status"] == "reveal":
        q = QUESTIONS[state["current_q"]]
        st.subheader(f"Question {state['current_q']+1} — Answer")
        st.markdown(f'<div class="qtext">{q["q"]}</div>', unsafe_allow_html=True)
        for i, opt in enumerate(q["options"]):
            marker = "✅ " if i == q["correct_index"] else "▫️ "
            st.write(f"{marker}{chr(65+i)}. {opt}")
        st.write("---")
        st.write("**Leaderboard so far**")
        for i, row in enumerate(leaderboard(state)[:8]):
            st.markdown(f'<div class="board-row"><span class="rank">#{i+1}</span><span>{row["name"]}</span><span class="score">{row["score"]}</span></div>', unsafe_allow_html=True)
        if time.time() - state["reveal_started_at"] >= REVEAL_SECONDS:
            nxt = state["current_q"] + 1
            if nxt >= len(QUESTIONS):
                state["status"] = "finished"
            else:
                state["current_q"] = nxt
                state["status"] = "active"
                state["question_started_at"] = time.time()
            st.rerun()

    elif state["status"] == "finished":
        st.title("🏁 Quiz complete")
        st.write(f"{len(QUESTIONS)} questions · {len(state['roster'])} students")
        for i, row in enumerate(leaderboard(state)):
            st.markdown(f'<div class="board-row"><span class="rank">#{i+1}</span><span>{row["name"]}</span><span class="score">{row["score"]}</span></div>', unsafe_allow_html=True)
        if st.button("Start a new session"):
            reset_quiz(state)
            st.rerun()


# ============================================================
# STUDENT VIEW
# ============================================================
def render_student():
    if "pid" not in st.session_state:
        st.session_state.pid = str(uuid.uuid4())
    if "joined" not in st.session_state:
        st.session_state.joined = False
    if "answered_q" not in st.session_state:
        st.session_state.answered_q = -1
    if "last_choice" not in st.session_state:
        st.session_state.last_choice = None

    pid = st.session_state.pid

    if not st.session_state.joined:
        st.markdown('<div class="eyebrow">BIOPHARMACEUTICS LIVE QUIZ</div>', unsafe_allow_html=True)
        st.title("Join the quiz")
        name = st.text_input("Your name", max_chars=24)
        if st.button("Join quiz", type="primary"):
            if name.strip():
                state["roster"][pid] = name.strip()
                st.session_state.joined = True
                st.rerun()
            else:
                st.warning("Enter your name first.")
        return

    st_autorefresh(interval=700, key="student_autorefresh")
    name = state["roster"].get(pid, "you")

    if state["status"] == "lobby":
        st.title(f"Hi {name} 👋")
        st.write("Waiting for the host to start the quiz...")
        return

    if state["status"] == "finished":
        board = leaderboard(state)
        my_rank = next((i+1 for i, r in enumerate(board) if r["name"] == name), None)
        my_score = next((r["score"] for r in board if r["name"] == name), 0)
        st.title(f"You finished #{my_rank or '-'} with {my_score} points")
        for i, row in enumerate(board):
            st.markdown(f'<div class="board-row"><span class="rank">#{i+1}</span><span>{row["name"]}</span><span class="score">{row["score"]}</span></div>', unsafe_allow_html=True)
        return

    q_idx = state["current_q"]
    q = QUESTIONS[q_idx]

    if state["status"] == "active":
        elapsed = time.time() - state["question_started_at"]
        remaining = max(0, q["time"] - elapsed)
        already_answered = state["answers"].get(q_idx, {}).get(pid) is not None

        st.subheader(f"Question {q_idx+1} of {len(QUESTIONS)}")
        st.progress(min(1.0, remaining / q["time"]))
        st.markdown(f"### ⏱ {int(remaining)+1}s")
        st.markdown(f'<div class="qtext">{q["q"]}</div>', unsafe_allow_html=True)

        if already_answered:
            st.info("Answer locked — waiting for the round to end.")
        elif remaining <= 0:
            # timed out without answering — record as a zero-score miss
            state["answers"].setdefault(q_idx, {})[pid] = {"choice": -1, "correct": False, "elapsed": q["time"], "score": 0}
            st.session_state.answered_q = q_idx
            st.session_state.last_choice = -1
            st.rerun()
        else:
            cols = st.columns(2)
            for i, opt in enumerate(q["options"]):
                if cols[i % 2].button(f"{chr(65+i)}. {opt}", key=f"opt_{q_idx}_{i}", use_container_width=True):
                    e = min(q["time"], time.time() - state["question_started_at"])
                    correct = (i == q["correct_index"])
                    score = compute_score(correct, e, q["time"])
                    state["answers"].setdefault(q_idx, {})[pid] = {"choice": i, "correct": correct, "elapsed": e, "score": score}
                    st.session_state.answered_q = q_idx
                    st.session_state.last_choice = i
                    st.rerun()

    elif state["status"] == "reveal":
        my_answer = state["answers"].get(q_idx, {}).get(pid)
        correct = my_answer["correct"] if my_answer else False
        score = my_answer["score"] if my_answer else 0
        st.title("Correct! ✅" if correct else "Not quite ❌")
        st.write(f"You scored **{score}** points this round")
        for i, opt in enumerate(q["options"]):
            marker = "✅ " if i == q["correct_index"] else "▫️ "
            st.write(f"{marker}{chr(65+i)}. {opt}")


def _base_url():
    # Best-effort guess for display purposes only; the actual link is whatever
    # this app's deployed Streamlit Cloud URL is.
    return "https://<your-app-name>.streamlit.app"


if role == "student":
    render_student()
else:
    render_host()
