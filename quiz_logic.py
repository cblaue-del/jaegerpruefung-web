#!/usr/bin/env python3
"""GUI-unabhängige Quiz-Logik (Reihenfolge, Auswertung, Ergebnisliste).

Getrennt von der Tkinter-Oberfläche, damit sie sich ohne Display testen lässt.
"""
import random


def _normalisiere_text(s):
    """Für den Vergleich eingetippter Antworten: Groß-/Kleinschreibung und
    überflüssige/mehrfache Leerzeichen spielen keine Rolle."""
    return " ".join((s or "").strip().casefold().split())


class QuizSession:
    def __init__(self, questions, shuffle=True, rng=None):
        self.questions = list(questions)
        if shuffle:
            (rng or random).shuffle(self.questions)
        self.index = 0
        self.results = []

    @property
    def current(self):
        if self.finished:
            raise IndexError("Quiz ist bereits beendet.")
        return self.questions[self.index]

    @property
    def total(self):
        return len(self.questions)

    @property
    def position(self):
        """1-basierte Nummer der aktuellen Frage."""
        return self.index + 1

    @staticmethod
    def correct_letters(question):
        return {a["letter"] for a in question["antworten"] if a["is_correct"]}

    @staticmethod
    def correct_texts(question):
        """Text der richtigen Antwort(en) – nützlich für die Auswertung, da die
        Anzeigereihenfolge (und damit die sichtbare Beschriftung A–D) bei jeder
        Anzeige der Frage neu gemischt wird und die Buchstaben allein daher
        nicht mehr eindeutig auf die Anzeige verweisen."""
        return [a["text"] for a in question["antworten"] if a["is_correct"]]

    def is_multi(self, question=None):
        question = question or self.current
        return len(self.correct_letters(question)) > 1

    @staticmethod
    def is_text(question):
        """True, wenn die Antwort eingetippt statt ausgewählt wird
        (Fragetyp 'Text' in der Excel-Vorlage)."""
        return question.get("typ") == "text"

    def check(self, selected_letters):
        """Prüft die Auswahl für die aktuelle (Multiple-Choice-)Frage. Exaktes
        Match nötig (alle richtigen ausgewählt, keine falschen). Rückgabe: bool."""
        q = self.current
        correct = self.correct_letters(q)
        selected = set(selected_letters)
        is_correct = selected == correct

        self.results.append(
            {
                "question_id": q["id"],
                "frage": q["frage"],
                "erlaeuterung": q["erlaeuterung"],
                "correct": is_correct,
                "selected": selected,
                "correct_letters": correct,
                "correct_texts": self.correct_texts(q),
            }
        )
        return is_correct

    def check_text(self, eingabe):
        """Prüft eine eingetippte Antwort für die aktuelle Text-Frage. Richtig,
        wenn sie (nach Normalisierung) mit einer der als richtig markierten
        Antworten übereinstimmt. Rückgabe: bool."""
        q = self.current
        akzeptiert = {_normalisiere_text(t) for t in self.correct_texts(q)}
        eingabe_norm = _normalisiere_text(eingabe)
        is_correct = bool(eingabe_norm) and eingabe_norm in akzeptiert

        self.results.append(
            {
                "question_id": q["id"],
                "frage": q["frage"],
                "erlaeuterung": q["erlaeuterung"],
                "correct": is_correct,
                "selected": {eingabe},
                "correct_letters": self.correct_letters(q),
                "correct_texts": self.correct_texts(q),
            }
        )
        return is_correct

    def advance(self):
        self.index += 1

    @property
    def finished(self):
        return self.index >= self.total

    @property
    def score(self):
        richtig = sum(1 for r in self.results if r["correct"])
        return richtig, len(self.results)

    @property
    def wrong_results(self):
        return [r for r in self.results if not r["correct"]]
