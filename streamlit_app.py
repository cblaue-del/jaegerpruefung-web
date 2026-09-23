#!/usr/bin/env python3
"""Jägerprüfung – Fragentrainer (Web-Version)

Browserbasierte Variante des Fragentrainers, gebaut mit Streamlit.
Nutzt exakt dieselbe Quiz-Logik und Datenbankanbindung wie die
Windows-Version (quiz_logic.py, db.py) – hier nur mit einer
Weboberfläche statt Tkinter.

Die Fragendatenbank (jagdfragen.db, Ordner "bilder") wird zentral
gepflegt und zusammen mit dieser Datei ausgeliefert. Es gibt in dieser
Version bewusst KEINEN Import-Dialog für Endnutzer – die Fragen kommen
ausschließlich aus der mitgelieferten Datenbank, aktualisiert wird sie,
indem eine neue Version davon ins Repository hochgeladen wird (siehe
README_WEB.md).

Es gibt außerdem bewusst kein Login und keinen dauerhaften Fortschritt:
jede Sitzung im Browser ist unabhängig, die Auswertung lebt nur so
lange, wie die Seite geöffnet ist. Das hält die App einfach und kommt
ohne Nutzerverwaltung aus.
"""
import base64
import os
import random
import sqlite3

import streamlit as st
from PIL import Image

import db
from quiz_logic import QuizSession

APP_TITLE = "Jägerprüfung – Fragentrainer"

BILD_MAX_BREITE = 640
BILD_MAX_HOEHE = 300

COLOR_BG = "#f5f3ee"
COLOR_ACCENT = "#2F5233"

HIER = os.path.dirname(os.path.abspath(__file__))


# --------------------------------------------------------------- Pfade
def db_path():
    return os.path.join(HIER, db.DB_FILENAME)


def bilder_dir():
    return os.path.join(HIER, "bilder")


def audio_dir():
    return os.path.join(HIER, "audio")


def resource_path(name):
    return os.path.join(HIER, name)


# ------------------------------------------------------ Datenbank/Cache
@st.cache_resource
def get_connection():
    """Eine gemeinsame, nur lesend genutzte Verbindung für alle Browser-
    Sitzungen. check_same_thread=False ist hier unbedenklich, da diese
    Web-Version nie in die Datenbank schreibt (kein Login/Fortschritt,
    siehe Moduldoku)."""
    conn = sqlite3.connect(db_path(), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    db.init_db(conn)
    return conn


@st.cache_data
def logo_base64(dateiname):
    pfad = resource_path(dateiname)
    try:
        with open(pfad, "rb") as f:
            return base64.b64encode(f.read()).decode("ascii")
    except OSError:
        return ""


def page_icon():
    pfad = resource_path("logo_icon.png")
    if os.path.isfile(pfad):
        return Image.open(pfad)
    return "🦌"


# ---------------------------------------------------------- Session-State
def init_state():
    ss = st.session_state
    ss.setdefault("screen", "start")
    ss.setdefault("session", None)
    ss.setdefault("revealed", False)
    ss.setdefault("checked", False)
    ss.setdefault("last_correct", None)
    ss.setdefault("selected_letters", set())
    ss.setdefault("selected_text", "")
    ss.setdefault("pruef_hinweis", None)
    ss.setdefault("answer_order", {})  # question_id -> gemischte Antwortliste


def _reset_question_state():
    st.session_state.revealed = False
    st.session_state.checked = False
    st.session_state.last_correct = None
    st.session_state.selected_letters = set()
    st.session_state.selected_text = ""
    st.session_state.pruef_hinweis = None


# ------------------------------------------------------------- Kopfzeile
def render_header():
    logo = logo_base64("logo_header.png")
    logo_html = (
        f'<img src="data:image/png;base64,{logo}" style="height:42px;" />'
        if logo
        else ""
    )
    st.markdown(
        f"""
        <div style="background-color:{COLOR_ACCENT};padding:14px 20px;
                    border-radius:8px;display:flex;align-items:center;
                    gap:14px;margin-bottom:24px;">
          {logo_html}
          <span style="color:white;font-size:21px;font-weight:600;">
            {APP_TITLE}
          </span>
        </div>
        """,
        unsafe_allow_html=True,
    )


# --------------------------------------------------------- Startbildschirm
def show_start_screen(conn):
    render_header()
    st.subheader("Bereit zum Üben?")

    total = db.question_count(conn)
    if total == 0:
        st.warning(
            "Es sind noch keine Fragen in der Datenbank hinterlegt. "
            "Bitte an den Betreiber der Seite wenden."
        )
        return

    st.caption(f"{total} Fragen in der Datenbank")

    kategorien = ["Alle"] + db.list_categories(conn)
    kategorie = st.selectbox("Kategorie", kategorien, key="start_kategorie")

    kat_filter = None if kategorie == "Alle" else kategorie
    anzahl_in_kategorie = len(db.get_questions(conn, kategorie=kat_filter))
    st.caption(f"{anzahl_in_kategorie} Fragen in dieser Auswahl")

    shuffle = st.checkbox("Zufällige Reihenfolge", value=True, key="start_shuffle")
    limit = st.number_input(
        "Maximale Anzahl Fragen (0 = alle)",
        min_value=0,
        max_value=total,
        value=0,
        step=1,
        key="start_limit",
        help="Für eine kurze Runde die Anzahl begrenzen – es wird dann eine "
        "zufällige Auswahl aus der gewählten Kategorie gezogen.",
    )

    st.button("Start", type="primary", on_click=_start_quiz, args=(conn, kategorie, shuffle, limit))


def _start_quiz(conn, kategorie, shuffle, limit):
    kat = None if kategorie == "Alle" else kategorie
    fragen = db.get_questions(conn, kategorie=kat)
    if not fragen:
        st.session_state.pruef_hinweis = None
        st.session_state.screen = "start"
        st.toast("Für diese Auswahl sind keine Fragen vorhanden.")
        return
    if limit and 0 < limit < len(fragen):
        fragen = random.sample(fragen, limit)
        if not shuffle:
            fragen.sort(key=lambda f: f["id"])
    st.session_state.session = QuizSession(fragen, shuffle=shuffle)
    st.session_state.answer_order = {}
    st.session_state.screen = "quiz"
    _reset_question_state()


def _neue_runde():
    st.session_state.screen = "start"
    st.session_state.session = None
    _reset_question_state()


def _nochmal_falsche(fragen):
    st.session_state.session = QuizSession(fragen, shuffle=True)
    st.session_state.answer_order = {}
    st.session_state.screen = "quiz"
    _reset_question_state()


# ----------------------------------------------------------- Fragebildschirm
def _zeige_bild(dateiname):
    pfad = os.path.join(bilder_dir(), dateiname)
    if not os.path.isfile(pfad):
        st.caption(f"(Bild nicht gefunden: {dateiname})")
        return
    try:
        img = Image.open(pfad)
        img.thumbnail((BILD_MAX_BREITE, BILD_MAX_HOEHE), Image.LANCZOS)
        st.image(img)
    except Exception:
        st.caption(f"(Bild konnte nicht geladen werden: {dateiname})")


def _spiele_audio(dateiname):
    pfad = os.path.join(audio_dir(), dateiname)
    if not os.path.isfile(pfad):
        st.caption(f"(Audio nicht gefunden: {dateiname})")
        return
    st.audio(pfad, format="audio/mp3")


def _antwort_label(position, antwort, correct_letters, selected_letters, checked):
    letter = antwort["letter"]
    label = f"{chr(65 + position)}) {antwort['text']}"
    if not checked:
        return label
    ist_richtig = letter in correct_letters
    ist_ausgewaehlt = letter in selected_letters
    if ist_richtig and ist_ausgewaehlt:
        return label + "  ✅"
    if ist_richtig and not ist_ausgewaehlt:
        return label + "  ⚠️ (richtige Antwort)"
    if not ist_richtig and ist_ausgewaehlt:
        return label + "  ❌"
    return label


def _reveal():
    st.session_state.revealed = True


def _pruefe_antwort():
    session = st.session_state.session
    q = session.current
    qid = q["id"]
    ist_text = QuizSession.is_text(q)

    if ist_text:
        eingabe = st.session_state.get(f"text_{qid}", "").strip()
        if not eingabe:
            st.session_state.pruef_hinweis = "Bitte gib eine Antwort ein."
            return
        is_correct = session.check_text(eingabe)
        st.session_state.selected_text = eingabe
    else:
        if session.is_multi(q):
            selected = {
                a["letter"]
                for a in q["antworten"]
                if st.session_state.get(f"cb_{qid}_{a['letter']}", False)
            }
        else:
            val = st.session_state.get(f"radio_{qid}")
            selected = {val} if val else set()
        if not selected:
            st.session_state.pruef_hinweis = "Bitte wähle mindestens eine Antwort aus."
            return
        is_correct = session.check(selected)
        st.session_state.selected_letters = selected

    st.session_state.checked = True
    st.session_state.last_correct = is_correct
    st.session_state.pruef_hinweis = None


def _weiter():
    session = st.session_state.session
    session.advance()
    if session.finished:
        st.session_state.screen = "result"
    _reset_question_state()


def show_question_screen(conn):
    render_header()
    session = st.session_state.session
    q = session.current
    qid = q["id"]

    kopf_links, kopf_rechts = st.columns([3, 1])
    with kopf_links:
        st.caption(f"Frage {session.position} von {session.total}")
    with kopf_rechts:
        if q["kategorie"]:
            st.caption(q["kategorie"])

    st.markdown(f"### {q['frage']}")

    if q.get("bild"):
        _zeige_bild(q["bild"])

    if q.get("audio"):
        _spiele_audio(q["audio"])

    if not st.session_state.revealed:
        st.button("Antworten anzeigen", type="primary", on_click=_reveal)
        return

    ist_text = QuizSession.is_text(q)
    multi = session.is_multi(q)

    if ist_text:
        st.caption("(Antwort bitte eintippen)")
    elif multi:
        st.caption("(Mehrere Antworten können richtig sein)")

    if qid not in st.session_state.answer_order:
        geordnet = list(q["antworten"])
        random.shuffle(geordnet)
        st.session_state.answer_order[qid] = geordnet
    antworten = st.session_state.answer_order[qid]

    checked = st.session_state.checked
    correct_letters = QuizSession.correct_letters(q)
    selected_letters = st.session_state.selected_letters

    if ist_text:
        st.text_input("Deine Antwort", key=f"text_{qid}", disabled=checked)
    elif multi:
        for pos, antwort in enumerate(antworten):
            letter = antwort["letter"]
            label = _antwort_label(pos, antwort, correct_letters, selected_letters, checked)
            st.checkbox(label, key=f"cb_{qid}_{letter}", disabled=checked)
    else:
        optionen = [a["letter"] for a in antworten]
        beschriftung = {
            a["letter"]: _antwort_label(i, a, correct_letters, selected_letters, checked)
            for i, a in enumerate(antworten)
        }
        st.radio(
            "Antwort wählen",
            optionen,
            index=None,
            format_func=lambda l: beschriftung[l],
            key=f"radio_{qid}",
            disabled=checked,
            label_visibility="collapsed",
        )

    if st.session_state.pruef_hinweis:
        st.warning(st.session_state.pruef_hinweis)

    if checked:
        if st.session_state.last_correct:
            st.success("Richtig!")
        else:
            st.error("Leider nicht ganz richtig.")
            if ist_text:
                richtige = ", ".join(QuizSession.correct_texts(q))
                st.markdown(f"**Richtige Antwort:** {richtige}")

    with st.expander("Erläuterung anzeigen"):
        st.write(q["erlaeuterung"] or "Keine Erläuterung hinterlegt.")

    knopf_links, knopf_mitte, knopf_rechts = st.columns([1, 1, 1])
    with knopf_links:
        st.button("Antwort prüfen", type="primary", disabled=checked, on_click=_pruefe_antwort)
    with knopf_rechts:
        st.button("Weiter →", disabled=not checked, on_click=_weiter)


# ------------------------------------------------------------- Ergebnis
def show_result_screen(conn):
    render_header()
    session = st.session_state.session
    richtig, gesamt = session.score

    st.subheader("Ergebnis")
    quote = (
        f"{richtig} von {gesamt} richtig ({round(100 * richtig / gesamt)} %)"
        if gesamt
        else "-"
    )
    st.markdown(f"#### {quote}")

    wrong = session.wrong_results
    if wrong:
        st.markdown("**Falsch beantwortete Fragen:**")
        for r in wrong:
            with st.container(border=True):
                st.markdown(f"• {r['frage']}")
                richtige_text = ", ".join(r["correct_texts"])
                st.markdown(f":green[Richtig gewesen wäre: {richtige_text}]")

        wrong_ids = {r["question_id"] for r in wrong}
        fragen_wiederholen = [f for f in session.questions if f["id"] in wrong_ids]
        st.button(
            "Falsch beantwortete nochmal üben",
            on_click=_nochmal_falsche,
            args=(fragen_wiederholen,),
        )
    else:
        st.success("Alle Fragen richtig beantwortet – stark!")

    st.button("Neue Runde", type="primary", on_click=_neue_runde)


# ------------------------------------------------------------------ main
def main():
    st.set_page_config(page_title=APP_TITLE, page_icon=page_icon(), layout="centered")
    init_state()
    conn = get_connection()

    screen = st.session_state.screen
    if screen == "quiz" and st.session_state.session is not None:
        show_question_screen(conn)
    elif screen == "result" and st.session_state.session is not None:
        show_result_screen(conn)
    else:
        show_start_screen(conn)


if __name__ == "__main__":
    main()
