"""Shared fixtures: synthetic PDFs (no private test papers are needed) and a fake AI transcriber."""
import io
import sys
from pathlib import Path

import pymupdf
import pytest
from PIL import Image, ImageDraw

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ai_transcriber import TranscriptionError  # noqa: E402

TITLE = "Test-99 (Synthetic Sample)"


def png_bytes(text, size=(300, 160)):
    img = Image.new("RGB", size, "white")
    ImageDraw.Draw(img).text((10, 10), text, fill="black")
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def _footer(page, kind, pno, total):
    page.insert_text((34, 815), f"{TITLE} · AEE · {kind}", fontsize=9)
    page.insert_text((512, 815), f"Page {pno} of {total}", fontsize=9)


def _option(page, y, text):
    page.draw_circle((44.25, y - 4.5), 4.2, color=(0, 0, 0), width=0.5)
    page.insert_text((63, y), text, fontsize=14)


def make_test_pdf(path):
    """3-page PDF in the Classplus layout: 2 questions (Q2 has an inline formula image),
    an answer key, and solutions (Solution 1 is a 'handwritten' image, Solution 2 printed text)."""
    doc = pymupdf.open()
    q = doc.new_page(width=595, height=842)
    q.insert_text((54, 66), "TEST", fontsize=9)
    q.insert_text((54, 90), TITLE, fontsize=9)
    q.insert_text((36, 232), "Section 1 : AEE - 2 Questions", fontsize=9)
    q.insert_text((36, 280), "1", fontsize=12)
    q.insert_text((54, 280), "What is 2+2?", fontsize=14)
    for y, t in zip((320, 350, 380, 410), ("3", "4", "5", "6")):
        _option(q, y, t)
    q.insert_text((518, 440), "Correct: +2", fontsize=9)
    q.insert_text((36, 490), "2", fontsize=12)
    q.insert_text((54, 490), "The value of x is", fontsize=14)
    q.insert_image(pymupdf.Rect(170, 476, 220, 496), stream=png_bytes("x=5", (100, 40)))
    for y, t in zip((530, 560, 590, 620), ("2", "5", "7", "9")):
        _option(q, y, t)
    q.insert_text((518, 650), "Correct: +2", fontsize=9)
    _footer(q, "Questions", 1, 3)

    a = doc.new_page(width=595, height=842)
    a.insert_text((36, 232), "Section 1 : AEE - 2 Questions", fontsize=9)
    a.insert_text((54, 270), "1", fontsize=9)
    a.insert_text((72, 270), "4", fontsize=9)
    a.insert_text((54, 320), "2", fontsize=9)
    a.insert_text((72, 320), "5", fontsize=9)
    _footer(a, "Answers", 2, 3)

    s = doc.new_page(width=595, height=842)
    s.insert_text((36, 232), "Section 1 : AEE - 2 Questions", fontsize=9)
    s.insert_text((36, 270), "1", fontsize=9)
    s.insert_image(pymupdf.Rect(37, 290, 337, 450), stream=png_bytes("2 + 2 = 4 (handwritten)"))
    s.insert_text((36, 480), "2", fontsize=9)
    s.insert_text((35, 505), "Because x = 5", fontsize=11)
    _footer(s, "Solutions", 3, 3)
    doc.save(path)
    return path


def make_generic_pdf(path):
    """3 pages without the test layout: printed text, a scanned image, and a page with no text/images."""
    doc = pymupdf.open()
    p1 = doc.new_page()
    p1.insert_text((72, 100), "Lecture notes: soil phases", fontsize=12)
    p2 = doc.new_page()
    p2.insert_image(pymupdf.Rect(72, 72, 472, 372), stream=png_bytes("handwritten page", (400, 300)))
    p3 = doc.new_page()
    p3.draw_line((72, 72), (300, 200), color=(0, 0, 1), width=2)  # pen strokes only
    doc.save(path)
    return path


class FakeTranscriber:
    """Stands in for ClaudeTranscriber: returns canned blocks per image file name."""

    model = "fake-model"

    def __init__(self, responses=None, default=None, fail=()):
        self.responses = responses or {}
        self.default = default if default is not None else [("p", "typed by fake AI")]
        self.fail = set(fail)
        self.calls = []

    def transcribe(self, path, context=""):
        name = Path(path).name
        self.calls.append((name, context))
        if name in self.fail:
            raise TranscriptionError("simulated API failure")
        blocks = self.responses.get(name, self.default)
        uncertain = ["second line smudged"] if name.startswith("solution") else []
        return blocks, uncertain, False


@pytest.fixture
def test_pdf(tmp_path):
    return make_test_pdf(tmp_path / "SYNTH-TEST.pdf")


@pytest.fixture
def generic_pdf(tmp_path):
    return make_generic_pdf(tmp_path / "NOTES.pdf")


@pytest.fixture
def quiet():
    return lambda *_: None
