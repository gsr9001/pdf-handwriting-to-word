"""Convert a Classplus-style test PDF (questions / answers / solutions) into a fully editable
Word document.

Strategy
  * Printed text is taken from the PDF's embedded text layer (PyMuPDF); superscripts kept.
  * The Questions / Answers / Solutions page ranges, the title and the question count are
    detected from the page footers and the first page.
  * Every embedded image (printed formula, table, handwritten solution or sketch) is replaced by
    an editable transcription - Word (OMML) equations, Word tables, text and text/arrow diagrams -
    taken from a per-test transcription module (manual visual transcription: no OCR engine is
    installed and Tesseract cannot read handwritten maths). No images are inserted.

Image content is resolved in this order:
  1. a manual transcription module: transcriptions_<PDF stem>.py next to this script, or
     --transcriptions FILE.py (for TEST-29 the original solution_transcriptions.py is used);
  2. with --ai, Claude's vision API (ai_transcriber.py) - results are cached as JSON in
     <out-dir>/ai_transcriptions/ and can be corrected by editing those files;
  3. otherwise the item is marked [UNCLEAR IN SOURCE] and listed in the review report.

PDFs that are not in the Classplus questions/answers/solutions layout are converted page by page
("generic" mode): printed text is kept, images and scanned/handwritten pages are transcribed.

Usage: python convert_pdf_to_word.py [input.pdf] [--out-dir DIR] [--transcriptions FILE.py]
                                     [--ai] [--model MODEL] [--overwrite] [--keep-images]
"""
import argparse
import importlib.util
import io
import re
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
import pymupdf
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt, RGBColor
from PIL import Image

from ai_transcriber import DEFAULT_MODEL, ClaudeTranscriber, TranscriptionError

FONT = "Calibri"
HERE = Path(__file__).resolve().parent

# set in configure()
PDF = OUT_DOCX = IMG_DIR = REVIEW = None
TITLE = SUBTITLE = ""
Q_PAGES = A_PAGES = S_PAGES = ()
EXPECTED_Q = 0
FORMULA_IMAGES, FIGURE_TABLES, SOLUTIONS, REVIEW_NOTES, REVIEW_GENERAL = {}, {}, {}, {}, []
UNCLEAR = "[UNCLEAR IN SOURCE]"
TRANSCRIPTION_FILE = None
GENERIC = False   # True when the PDF is not in the questions/answers/solutions layout
AI = None         # ClaudeTranscriber when --ai is used
KEEP_IMAGES = False   # True: place the original images in the document instead of typing them out
MAX_PICTURE_WIDTH = 6.0   # inches; the text column of the A4 page
LOG = print


class ConversionError(Exception):
    """The PDF cannot be converted (missing, unreadable, encrypted, empty)."""

SUP_ON, SUP_OFF = "\x01", "\x02"
review = []
arrow_noted = []


# ---------------------------------------------------------------- extraction
def save_image(doc, xref, path):
    info = doc.extract_image(xref)
    img = Image.open(io.BytesIO(info["image"]))
    if info.get("smask"):
        mask = Image.open(io.BytesIO(doc.extract_image(info["smask"])["image"])).convert("L")
        img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, "white")
        bg.paste(img, mask=mask.resize(img.size))
        img = bg
    img.convert("RGB").save(path, "PNG")
    return img.size


def page_elements(doc, pno, cutoff_text="Section 1", strip_footer=True):
    """Return (elements, circles) for a page, header/footer noise removed.
    element = dict(kind 't'|'i', x0,y0,x1,y1, text|xref|digest)"""
    page = doc[pno - 1]
    els = []
    for b in page.get_text("dict")["blocks"]:
        if b["type"] != 0:
            continue
        for ln in b["lines"]:
            # superscript spans (flag bit 0) are wrapped in SUP_ON/SUP_OFF markers
            txt = "".join(SUP_ON + s["text"] + SUP_OFF if s["flags"] & 1 and s["text"].strip() else s["text"]
                          for s in ln["spans"]).replace("\xa0", " ")
            if "" in txt:  # Wingdings right-arrow glyph (private-use code point)
                txt = txt.replace("", "→")
                if not arrow_noted:
                    arrow_noted.append(1)
                    review.append(f"Wingdings arrow glyph (U+F0E0) in the PDF text (first on p{pno}) was mapped to '→'. Verify.")
            txt = txt.strip()
            if txt:
                x0, y0, x1, y1 = ln["bbox"]
                els.append(dict(kind="t", x0=x0, y0=y0, x1=x1, y1=y1, text=txt))
    for im in page.get_image_info(xrefs=True):
        if im["xref"]:
            x0, y0, x1, y1 = im["bbox"]
            els.append(dict(kind="i", x0=x0, y0=y0, x1=x1, y1=y1, xref=im["xref"], digest=im["digest"]))
    cutoff = 0
    for e in els:
        if e["kind"] == "t" and e["text"].startswith(cutoff_text) and e["y0"] < 260:
            cutoff = e["y1"] + 2
    out = []
    n_correct = 0
    for e in els:
        if (strip_footer and e["y0"] >= 800) or e["y0"] < cutoff:
            continue
        if e["kind"] == "t" and re.fullmatch(r"Correct: \+\d+", e["text"]):
            n_correct += 1
            continue
        out.append(e)
    circles = [d["rect"].y0 + d["rect"].height / 2 for d in page.get_drawings()
               if d["type"] == "s" and 7 < d["rect"].width < 10 and 7 < d["rect"].height < 10 and d["rect"].x0 < 50]
    return sorted(out, key=lambda e: (e["y0"], e["x0"])), circles, n_correct


def group_rows(els):
    rows = []
    for e in els:
        if rows:
            r = rows[-1]
            ov = min(r["y1"], e["y1"]) - max(r["y0"], e["y0"])
            if ov > 0.3 * min(r["y1"] - r["y0"], e["y1"] - e["y0"]):
                r["els"].append(e)
                r["y0"], r["y1"] = min(r["y0"], e["y0"]), max(r["y1"], e["y1"])
                continue
        rows.append(dict(els=[e], y0=e["y0"], y1=e["y1"]))
    for r in rows:
        r["els"].sort(key=lambda e: e["x0"])
        r["x0"] = r["els"][0]["x0"]
    return rows


# ---------------------------------------------------------------- parsing
def split_number(row, xmax):
    """Return (number, row without the number) if the row starts with a question number at the left
    margin - either as its own text line ("7") or fused with the text ("7 A line preserving ...")."""
    first = row["els"][0]
    if first["kind"] != "t" or first["x0"] >= xmax:
        return None, row
    m = re.fullmatch(r"(\d+)(?:\s+(\S.*))?", first["text"])
    if not m:
        return None, row
    rest = row["els"][1:]
    if m.group(2):
        rest = [dict(first, text=m.group(2))] + rest
    row = dict(row, els=rest)
    if rest:
        row["x0"] = rest[0]["x0"]
    return int(m.group(1)), row


class Images:
    def __init__(self, doc):
        self.doc, self.cache = doc, {}

    def get(self, e, name):
        if e["xref"] in self.cache:
            return self.cache[e["xref"]]
        path = IMG_DIR / name
        px = save_image(self.doc, e["xref"], path)
        self.cache[e["xref"]] = (path, px)
        return self.cache[e["xref"]]


def row_runs(row, imgs, namer):
    runs = []
    for e in row["els"]:
        if e["kind"] == "t":
            if runs and runs[-1][0] == "t":
                runs[-1] = ("t", runs[-1][1] + " " + e["text"])
            else:
                runs.append(("t", e["text"]))
        else:
            path, _ = imgs.get(e, namer())
            runs.append(("i", path, e["x1"] - e["x0"], e["y1"] - e["y0"]))
    return runs


def lift_question_numbers(els):
    """Some PDFs print a standalone question number level with the LAST line of a wrapped stem.
    Move such a number up to the first line of that contiguous stem block."""
    texts = [e for e in els if e["kind"] == "t" and e["x0"] >= 45]
    for e in els:
        if not (e["kind"] == "t" and re.fullmatch(r"\d+", e["text"]) and e["x0"] < 45):
            continue
        cur = next((t for t in texts if min(t["y1"], e["y1"]) - max(t["y0"], e["y0"]) > 0.5 * (e["y1"] - e["y0"])), None)
        while cur is not None:
            above = [t for t in texts if 0 <= cur["y0"] - t["y1"] < 4 and abs(t["x0"] - cur["x0"]) < 3]
            if not above:
                break
            cur = max(above, key=lambda t: t["y1"])
        if cur is not None and cur["y0"] < e["y0"] - 4:
            h = e["y1"] - e["y0"]
            e["y0"], e["y1"] = cur["y0"] + 1, cur["y0"] + 1 + h
    return sorted(els, key=lambda e: (e["y0"], e["x0"]))


def parse_questions(doc, imgs):
    qs, cur, last, correct = {}, None, None, 0
    counters = {}
    for pno in Q_PAGES:
        els, circles, nc = page_elements(doc, pno)
        els = lift_question_numbers(els)
        correct += nc
        last = None  # never merge lines across pages
        for row in group_rows(els):
            num, row = split_number(row, 50)
            if num is not None:
                cur = num
                qs[cur] = dict(paras=[], pages=[pno])
                counters[cur] = 0
                if not row["els"]:
                    continue
            if cur is None:
                review.append(f"p{pno}: content before first question number ignored: {[e.get('text') for e in row['els']]}")
                continue
            ymid = (row["y0"] + row["y1"]) / 2
            is_opt = any(abs(c - ymid) < 9 for c in circles)
            if is_opt and sum(p["kind"] == "option" for p in qs[cur]["paras"]) >= 4:
                cur += 1
                qs[cur] = dict(paras=[], pages=[pno], no_stem=True)
                counters[cur] = 0
                last = None
                review.append(f"Q{cur}: question number AND stem are absent from the PDF (p{pno} starts directly with "
                              f"its options, following Q{cur - 1}); number inferred from sequence, stem text unavailable. "
                              f"Manual review needed.")
            elif (not is_opt and qs[cur]["paras"]
                    and any(p["kind"] == "option" for p in qs[cur]["paras"])
                    and qs[cur]["paras"][-1]["kind"] == "option"
                    and not (qs[cur]["paras"][-1]["x0"] and row["y0"] - qs[cur]["paras"][-1]["y1"] < 5 and last is qs[cur]["paras"][-1])):
                cur += 1
                qs[cur] = dict(paras=[], pages=[pno])
                counters[cur] = 0
                last = None
                review.append(f"Q{cur}: question number is missing from the PDF text layer (p{pno}); number inferred from "
                              f"sequence (follows Q{cur - 1}). Please confirm against the original.")
            q = qs[cur]
            if pno not in q["pages"]:
                q["pages"].append(pno)
            single_big = len(row["els"]) == 1 and row["els"][0]["kind"] == "i" and \
                (row["els"][0]["x1"] - row["els"][0]["x0"]) > 150

            def namer(cur=cur, is_opt=is_opt, big=single_big):
                counters[cur] += 1
                k = counters[cur]
                if big:
                    return f"question_q{cur:02d}_table.png" if k == 1 else f"question_q{cur:02d}_figure{k}.png"
                return f"question_q{cur:02d}_{'option' if is_opt else 'formula'}{k}.png"

            runs = row_runs(row, imgs, namer)
            kind = "option" if is_opt else "figure" if single_big else "body"
            prev = q["paras"][-1] if q["paras"] else None
            if (last is prev and prev is not None and prev["kind"] == kind and kind in ("body", "option")
                    and abs(prev["x0"] - row["x0"]) < 3 and row["y0"] - prev["y1"] < 5 and kind == "body"
                    or (last is prev and prev is not None and prev["kind"] == "option" and not is_opt
                        and row["y0"] - prev["y1"] < 5 and not single_big)):
                if runs[0][0] == "t" and prev["runs"][-1][0] == "t":
                    prev["runs"][-1] = ("t", prev["runs"][-1][1] + " " + runs[0][1])
                    runs = runs[1:]
                prev["runs"] += runs
                prev["y1"] = row["y1"]
            else:
                p = dict(kind=kind, runs=runs, x0=row["x0"], y1=row["y1"])
                q["paras"].append(p)
                last = p
    return qs, correct


def parse_answers(doc, imgs):
    ans = {}
    for pno in A_PAGES:
        els, _, _ = page_elements(doc, pno)
        for row in group_rows(els):
            n, rest = split_number(row, 70)
            if n is not None:
                ans[n] = row_runs(rest, imgs, lambda n=n: f"answer_q{n:02d}.png")
            else:
                review.append(f"p{pno}: unattached answer-key row {[e.get('text') for e in row['els']]}")
    return ans


def parse_solutions(doc, imgs):
    sols, cur, seen = {}, None, {}
    for pno in S_PAGES:
        els, _, _ = page_elements(doc, pno)
        for e in els:
            if e["kind"] == "t" and re.fullmatch(r"\d+", e["text"]) and e["x0"] < 50:
                cur = int(e["text"])
                sols.setdefault(cur, [])
            elif e["kind"] == "i":
                if cur is None:
                    review.append(f"p{pno}: solution image before any solution number")
                    continue
                if (cur, e["digest"]) in seen:
                    review.append(f"Solution {cur}: duplicate image on p{pno} skipped")
                    continue
                seen[(cur, e["digest"])] = True
                k = sum(1 for x in sols[cur] if x[0] == "img") + 1
                path, px = imgs.get(e, f"solution_q{cur:02d}_{k}.png")
                sols[cur].append(("img", path, px, pno))
            elif cur is None:
                review.append(f"p{pno}: solution-page text before any solution number {e['text']!r}")
            else:
                sols[cur].append(("text", e["text"], pno))
    return sols


# ---------------------------------------------------------------- docx
def set_font(style, size, bold=None, color=None):
    style.font.name = FONT
    style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color:
        style.font.color.rgb = color


def add_field(par, instr):
    for t, txt in (("begin", None), (None, instr), ("end", None)):
        r = par.add_run()
        if t:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), t)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = txt
        r._r.append(el)


# ---- equations: tiny LaTeX-like markup -> Office Math (OMML)
def parse_math(s, i=0):
    """Parse markup into nodes until a closing '}' or end. Returns (nodes, index)."""
    nodes = []

    def group(j):
        if j < len(s) and s[j] == "{":
            return parse_math(s, j + 1)
        return [("r", s[j], False)], j + 1

    while i < len(s):
        c = s[i]
        if c == "}":
            return nodes, i + 1
        if s.startswith("\\frac", i):
            num, i = group(i + 5)
            den, i = group(i)
            nodes.append(("f", num, den))
        elif s.startswith("\\box", i):
            body, i = group(i + 4)
            nodes.append(("box", body))
        elif s.startswith("\\text", i):
            j = s.index("}", i)
            nodes.append(("r", s[i + 6:j], True))
            i = j + 1
        elif c in "_^":
            arg, i = group(i + 1)
            base = nodes.pop() if nodes else ("r", "", False)
            if base[0] == "r" and not base[2] and len(base[1]) > 1:
                nodes.append(("r", base[1][:-1], False))
                base = ("r", base[1][-1], False)
            if c == "^" and base[0] == "sub":
                nodes.append(("subsup", base[1], base[2], arg))
            else:
                nodes.append(("sub" if c == "_" else "sup", [base], arg))
        elif c == "{":
            body, i = parse_math(s, i + 1)
            nodes.extend(body)
        else:
            if nodes and nodes[-1][0] == "r" and not nodes[-1][2]:
                nodes[-1] = ("r", nodes[-1][1] + c, False)
            else:
                nodes.append(("r", c, False))
            i += 1
    return nodes, i


def _m(tag, *children):
    el = OxmlElement("m:" + tag)
    for ch in children:
        el.append(ch)
    return el


def omml_nodes(nodes):
    out = []
    for n in nodes:
        if n[0] == "r":
            r = _m("r")
            rpr = _m("rPr")
            if n[2]:
                rpr.append(_m("nor"))
            else:
                sty = _m("sty")
                sty.set(qn("m:val"), "p")
                rpr.append(sty)
            r.append(rpr)
            t = _m("t")
            t.set(qn("xml:space"), "preserve")
            t.text = n[1]
            r.append(t)
            out.append(r)
        elif n[0] == "f":
            out.append(_m("f", _m("num", *omml_nodes(n[1])), _m("den", *omml_nodes(n[2]))))
        elif n[0] == "box":
            out.append(_m("borderBox", _m("e", *omml_nodes(n[1]))))
        elif n[0] in ("sub", "sup"):
            out.append(_m("s" + n[0].capitalize(), _m("e", *omml_nodes(n[1])), _m(n[0], *omml_nodes(n[2]))))
        elif n[0] == "subsup":
            out.append(_m("sSubSup", _m("e", *omml_nodes(n[1])), _m("sub", *omml_nodes(n[2])),
                          _m("sup", *omml_nodes(n[3]))))
    return out


def omath(markup):
    return _m("oMath", *omml_nodes(parse_math(markup)[0]))


def add_display_eq(doc, markup):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.space_after = Pt(5)
    jc = _m("jc")
    jc.set(qn("m:val"), "left")
    p._p.append(_m("oMathPara", _m("oMathParaPr", jc), omath(markup)))
    return p


def add_markup_text(par, text, bold=False, size=None):
    """Plain text with $inline math$ segments."""
    for k, seg in enumerate(text.split("$")):
        if k % 2:
            par._p.append(omath(seg))
        elif seg:
            add_text(par, seg, bold=bold, size=size)


def add_text(par, text, bold=False, size=None):
    """Add text, rendering SUP_ON..SUP_OFF segments (from the PDF) as superscript."""
    for k, seg in enumerate(re.split(f"[{SUP_ON}{SUP_OFF}]", text)):
        if seg:
            run = par.add_run(seg)
            run.bold = bold
            run.font.superscript = bool(k % 2)
            if size:
                run.font.size = Pt(size)


def transcribe_image(path, context):
    """Blocks for an image that has no manual transcription: Claude (with --ai) or None."""
    global AI
    if AI is None:
        return None
    try:
        blocks, uncertain, cached = AI.transcribe(path, context)
    except TranscriptionError as exc:
        if getattr(exc, "fatal", False):  # e.g. no credentials: stop calling the API for this run
            LOG(f"AI transcription disabled: {exc}")
            review.append(f"AI transcription stopped at {context}: {exc}. Remaining images are marked {UNCLEAR}.")
            AI = None
            return None
        review.append(f"{context}: AI transcription failed ({exc}) — marked {UNCLEAR}.")
        return None
    except Exception as exc:  # e.g. network failure inside a custom client - never abort the whole document
        review.append(f"{context}: AI transcription failed ({type(exc).__name__}: {exc}) — marked {UNCLEAR}.")
        return None
    ai_transcribed.append(path.name)
    review.append(f"{context}: transcribed by AI ({AI.model}{', from cache' if cached else ''}) — "
                  "proof-read against the PDF.")
    review.extend(f"{context}: AI flagged as uncertain — {u}" for u in uncertain)
    return blocks


def add_picture(par, path, width_in):
    """Place the original image, unmodified, in `par` (never wider than the text column)."""
    par.add_run().add_picture(str(path), width=Inches(min(width_in, MAX_PICTURE_WIDTH)))


def add_picture_block(doc, path, width_in):
    """The original image on its own centred line."""
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_after = Pt(6)
    par.paragraph_format.keep_together = True
    add_picture(par, path, width_in)


def add_runs(par, runs, scale=0.85, bold_text=False, context=""):
    for r in runs:
        if r[0] == "t":
            add_text(par, r[1] + " ", bold=bold_text)
            continue
        if r[1].name in FORMULA_IMAGES:
            used_formulas.add(r[1].name)
            par._p.append(omath(FORMULA_IMAGES[r[1].name]))
            par.add_run(" ")
            continue
        if KEEP_IMAGES:
            add_picture(par, r[1], r[2] / 72)
            par.add_run(" ")
            continue
        blocks = transcribe_image(r[1], f"{context} inline image {r[1].name}".strip())
        if blocks and all(k in ("eq", "p") for k, *_ in blocks):
            for kind, val in blocks:
                if kind == "eq":
                    par._p.append(omath(val))
                else:
                    add_markup_text(par, val)
                par.add_run(" ")
        else:
            if blocks:
                review.append(f"{context}: inline image {r[1].name} holds a table/diagram — not placed inline, "
                              f"marked {UNCLEAR}.")
            elif AI is None:
                review.append(f"{context}: no transcription for image {r[1].name} — marked {UNCLEAR}.")
            par.add_run(f"{UNCLEAR} ")


def add_table(doc, rows, widths=None, size=10, keep_with_next=False):
    tbl = doc.add_table(rows=0, cols=len(rows[0]))
    tbl.style = "Table Grid"
    for ri, cells in enumerate(rows):
        row = tbl.add_row()
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for ci, txt in enumerate(cells):
            c = row.cells[ci]
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.keep_with_next = keep_with_next or ri < len(rows) - 1
            add_markup_text(p, txt, bold=ri == 0, size=size)
            if ri == 0:
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:fill"), "D9E2F3")
                c._tc.get_or_add_tcPr().append(shd)
            if widths:
                c.width = Inches(widths[ci])
    # rows whose trailing cells are empty become full-width rows
    for ri, cells in enumerate(rows):
        if len(cells) > 1 and all(x == "" for x in cells[1:]):
            r = tbl.rows[ri]
            r.cells[0].merge(r.cells[-1])
    tbl.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    return tbl


def add_diagram(doc, rows):
    """Simple diagram as a borderless, centred grid of editable labels and arrow characters."""
    tbl = doc.add_table(rows=0, cols=len(rows[0]))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ri, cells in enumerate(rows):
        row = tbl.add_row()
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for ci, txt in enumerate(cells):
            c = row.cells[ci]
            c.width = Inches(6.0 / len(cells))
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.keep_with_next = ri < len(rows) - 1
            arrow = len(txt) <= 2
            centre = len(rows) == 3 and len(cells) == 3 and ri == 1 and ci == 1 and "•" not in txt
            if txt != "<<":
                add_markup_text(p, txt, bold=centre, size=16 if arrow else 10.5)
    for ri, cells in enumerate(rows):  # "<<" in a cell merges it into the cell on its left
        for ci in range(len(cells) - 1, 0, -1):
            if cells[ci] == "<<":
                tbl.rows[ri].cells[ci - 1].merge(tbl.rows[ri].cells[ci])
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return tbl


used_formulas = set()
ai_transcribed = []


def new_document():
    """A4 document with the shared styles, running header and page-number footer."""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, m, Inches(1))
    st = doc.styles
    set_font(st["Normal"], 11)
    st["Normal"].paragraph_format.line_spacing = 1.15
    st["Normal"].paragraph_format.space_after = Pt(6)
    set_font(st["Title"], 20, True, RGBColor(0, 0, 0))
    st["Title"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for n, sz in (("Heading 1", 16), ("Heading 2", 12)):
        set_font(st[n], sz, True, RGBColor(0x1F, 0x3A, 0x5F))
    st["Heading 1"].paragraph_format.space_before = Pt(0)
    st["Heading 1"].paragraph_format.space_after = Pt(10)
    st["Heading 2"].paragraph_format.space_before = Pt(10)
    st["Heading 2"].paragraph_format.space_after = Pt(4)

    hp = sec.header.paragraphs[0]
    hp.text = TITLE
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.runs[0].font.size = Pt(9)
    hp.runs[0].font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.add_run("Page ")
    add_field(fp, "PAGE")
    fp.add_run(" of ")
    add_field(fp, "NUMPAGES")

    doc.add_paragraph(TITLE, style="Title")
    if SUBTITLE:
        sub = doc.add_paragraph(SUBTITLE)
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return doc


def build_docx(qs, ans, sols):
    doc = new_document()

    # --- Section A
    doc.add_heading("1. Test Title and Questions", 1)
    labels = "ABCDEFGH"
    for n in sorted(qs):
        paras = qs[n]["paras"]
        opts_total = sum(1 for p in paras if p["kind"] == "option")
        first_body = True
        oi = 0
        if qs[n].get("no_stem"):
            par = doc.add_paragraph()
            par.paragraph_format.keep_with_next = True
            par.paragraph_format.space_before = Pt(8)
            par.add_run(f"Q{n}.  ").bold = True
            r = par.add_run("[Question text is missing from the source PDF — see conversion_review.txt]")
            r.italic = True
            r.font.color.rgb = RGBColor(0xC0, 0, 0)
            first_body = False
        for idx, p in enumerate(paras):
            par = doc.add_paragraph()
            pf = par.paragraph_format
            pf.keep_together = True
            pf.keep_with_next = idx < len(paras) - 1
            pf.space_after = Pt(4)
            if p["kind"] == "option":
                pf.left_indent = Inches(0.45)
                pf.first_line_indent = Inches(-0.3)
                par.add_run(f"({labels[oi]})  ")
                oi += 1
                add_runs(par, p["runs"], context=f"Q{n} (PDF page {', '.join(map(str, qs[n]['pages']))})")
            elif p["kind"] == "figure":
                _, path, w, h = p["runs"][0]
                if KEEP_IMAGES:
                    par._p.getparent().remove(par._p)
                    add_picture_block(doc, path, w / 72)
                elif path.name in FIGURE_TABLES:
                    par._p.getparent().remove(par._p)
                    rows = FIGURE_TABLES[path.name]
                    add_table(doc, rows, widths=(2.0, 4.0) if len(rows[0]) == 2 else None, size=11,
                              keep_with_next=True)
                    spacer = doc.add_paragraph()
                    spacer.paragraph_format.space_after = Pt(0)
                    spacer.paragraph_format.keep_with_next = True
                else:
                    blocks = transcribe_image(path, f"Q{n} figure {path.name}")
                    if blocks:
                        par._p.getparent().remove(par._p)
                        add_blocks(doc, blocks)
                    else:
                        if AI is None:
                            review.append(f"Q{n}: figure {path.name} has NO transcription — marked {UNCLEAR}.")
                        pf.left_indent = Inches(0.3)
                        par.add_run(f"{UNCLEAR} — figure not transcribed")
            else:
                if first_body:
                    pf.space_before = Pt(8)
                    r = par.add_run(f"Q{n}.  ")
                    r.bold = True
                    first_body = False
                elif p["x0"] > 55:
                    pf.left_indent = Inches(0.3)
                add_runs(par, p["runs"], context=f"Q{n} (PDF page {', '.join(map(str, qs[n]['pages']))})")
        if first_body:
            review.append(f"Q{n}: no stem text/body found")
        if opts_total != 4:
            review.append(f"Q{n}: {opts_total} options detected (expected 4) — verify against PDF pages {qs[n]['pages']}")

    # --- Section B
    doc.add_page_break()
    doc.add_heading("2. Answer Key", 1)
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c, txt in zip(tbl.rows[0].cells, ("Q. No.", "Answer")):
        c.text = ""
        r = c.paragraphs[0].add_run(txt)
        r.bold = True
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), "D9E2F3")
        c._tc.get_or_add_tcPr().append(shd)
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    trPr.append(th)
    for n in sorted(ans):
        row = tbl.add_row()
        cs = OxmlElement("w:cantSplit")
        row._tr.get_or_add_trPr().append(cs)
        row.cells[0].text = ""
        row.cells[0].paragraphs[0].add_run(str(n)).bold = True
        row.cells[1].text = ""
        add_runs(row.cells[1].paragraphs[0], ans[n], scale=0.8, context=f"Answer {n}")
        for c in row.cells:
            c.paragraphs[0].paragraph_format.space_after = Pt(0)
            c.paragraphs[0].paragraph_format.line_spacing = 1.0
    for r in tbl.rows:
        r.cells[0].width, r.cells[1].width = Inches(0.9), Inches(5.3)

    # --- Section C: printed solution text + typed transcriptions (no images are inserted)
    doc.add_page_break()
    doc.add_heading("3. Detailed Solutions — Fully Typed and Editable", 1)
    intro = doc.add_paragraph()
    if KEEP_IMAGES:
        r = intro.add_run(f"The solutions of the source PDF (pages {S_PAGES[0]}–{S_PAGES[-1]}) are given below. "
                          "Printed text is editable; handwritten, drawn and scanned content is shown as the "
                          "original image from the PDF.")
    else:
        r = intro.add_run(f"The solutions of the source PDF (pages {S_PAGES[0]}–{S_PAGES[-1]}) are given below as "
                          "editable text. Handwritten or image content is typed as text, Word equations and tables, "
                          "in the order written; spelling and arithmetic are reproduced as in the source. Unreadable "
                          f"parts are marked {UNCLEAR}.")
    r.italic = True
    r.font.size = Pt(9.5)
    for n in sorted(sols):
        h = doc.add_heading(f"Solution {n}", 2)
        h.paragraph_format.keep_with_next = True
        items = sols[n]
        pages = sorted({x[-1] for x in items})
        if not items:
            p = doc.add_paragraph()
            r = p.add_run("(No solution is given in the source PDF for this question.)")
            r.italic = True
            r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
            review.append(f"Solution {n}: number appears in the PDF but no solution follows it — nothing to transcribe.")
            continue
        if n in REVIEW_NOTES:
            review.append(f"Solution {n} (PDF page {', '.join(map(str, pages))}): {REVIEW_NOTES[n]}")
        blocks_done = False
        for item in items:
            if item[0] == "text":
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.space_after = Pt(5)
                add_text(p, item[1])
                continue
            if KEEP_IMAGES:
                add_picture_block(doc, item[1], item[2][0] / 120)  # item[2] = (width, height) in pixels
                continue
            if n in SOLUTIONS:  # one manual transcription covers all images of a solution
                if not blocks_done:
                    add_blocks(doc, SOLUTIONS[n])
                blocks_done = True
                continue
            blocks = transcribe_image(item[1], f"Solution {n} (PDF page {item[3]})")
            if not blocks:
                doc.add_paragraph(f"{UNCLEAR} — image content not transcribed.")
                if AI is None:
                    review.append(f"Solution {n} (PDF page {item[3]}): image has NO transcription.")
                continue
            add_blocks(doc, blocks)
    missing_f = set(FORMULA_IMAGES) - used_formulas
    if missing_f:
        review.append(f"Formula transcriptions not used (image not found): {sorted(missing_f)}")
    return doc


def build_generic_docx(doc_pdf, imgs):
    """Page-by-page conversion for PDFs without the questions/answers/solutions layout."""
    doc = new_document()
    for pno in range(1, len(doc_pdf) + 1):
        doc.add_heading(f"Page {pno}", 1)
        els, _, _ = page_elements(doc_pdf, pno, cutoff_text="\0", strip_footer=False)
        rows = group_rows(els)
        k = 0
        for row in rows:
            texts = [e["text"] for e in row["els"] if e["kind"] == "t"]
            if texts:
                p = doc.add_paragraph()
                add_text(p, " ".join(texts))
            for e in row["els"]:
                if e["kind"] != "i":
                    continue
                k += 1
                path, _ = imgs.get(e, f"page{pno:02d}_img{k}.png")
                if KEEP_IMAGES:
                    add_picture_block(doc, path, (e["x1"] - e["x0"]) / 72)
                    continue
                blocks = transcribe_image(path, f"Page {pno} image {k}")
                if blocks:
                    add_blocks(doc, blocks)
                else:
                    doc.add_paragraph(f"{UNCLEAR} — image content not transcribed.")
                    if AI is None:
                        review.append(f"Page {pno} image {k}: no transcription — marked {UNCLEAR}.")
        if not rows:
            # no text layer and no embedded images: a vector/ink page - transcribe a rendering of it
            path = IMG_DIR / f"page{pno:02d}_render.png"
            doc_pdf[pno - 1].get_pixmap(dpi=200).save(path)
            if KEEP_IMAGES:
                add_picture_block(doc, path, MAX_PICTURE_WIDTH)
                continue
            blocks = transcribe_image(path, f"Page {pno} (rendered page)")
            if blocks:
                add_blocks(doc, blocks)
            else:
                doc.add_paragraph(f"{UNCLEAR} — page content not transcribed.")
                if AI is None:
                    review.append(f"Page {pno}: no text layer — needs --ai or a manual transcription; "
                                  f"marked {UNCLEAR}.")
    return doc


# ---------------------------------------------------------------- main
def add_blocks(doc, blocks):
    for bi, (kind, *val) in enumerate(blocks):
        nxt = blocks[bi + 1][0] if bi + 1 < len(blocks) else None
        if kind == "eq":
            add_display_eq(doc, val[0])
        elif kind == "p":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.3)
            p.paragraph_format.space_after = Pt(5)
            p.paragraph_format.keep_with_next = nxt in ("table", "diagram")
            add_markup_text(p, val[0])
        elif kind == "table":
            add_table(doc, val[0])
            doc.add_paragraph().paragraph_format.space_after = Pt(0)
        elif kind == "diagram":
            add_diagram(doc, val[0])
        elif kind == "note":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.3)
            r = p.add_run("Transcriber’s note: " + val[0])
            r.italic = True
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x80, 0x60, 0x00)


def load_transcriptions(path):
    spec = importlib.util.spec_from_file_location("transcriptions", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def unique_path(path, overwrite):
    """`path`, or `name (2).ext`, `name (3).ext`... so existing outputs are never overwritten."""
    if overwrite or not path.exists():
        return path
    k = 2
    while path.with_name(f"{path.stem} ({k}){path.suffix}").exists():
        k += 1
    return path.with_name(f"{path.stem} ({k}){path.suffix}")


def configure(doc, out_dir, transcriptions, overwrite):
    """Detect the layout, page ranges, title and question count; pick output paths and transcriptions."""
    global OUT_DOCX, IMG_DIR, REVIEW, TITLE, SUBTITLE, Q_PAGES, A_PAGES, S_PAGES, EXPECTED_Q, GENERIC
    global FORMULA_IMAGES, FIGURE_TABLES, SOLUTIONS, REVIEW_NOTES, REVIEW_GENERAL, UNCLEAR, TRANSCRIPTION_FILE
    kinds = {}
    for i, page in enumerate(doc):
        foot = " ".join(page.get_text("text", clip=pymupdf.Rect(0, 795, page.rect.width, page.rect.height)).split())
        m = re.search(r"·\s*(Questions|Answers|Solutions)\b", foot)
        kinds.setdefault(m.group(1) if m else None, []).append(i + 1)
    GENERIC = None in kinds or "Questions" not in kinds
    Q_PAGES, A_PAGES, S_PAGES = kinds.get("Questions", []), kinds.get("Answers", []), kinds.get("Solutions", [])
    lines = [ln.strip() for ln in doc[0].get_text().splitlines() if ln.strip()]
    section, category = "", None
    if GENERIC:
        Q_PAGES = A_PAGES = S_PAGES = []
        pdf_title = (doc.metadata or {}).get("title") or PDF.stem
        EXPECTED_Q = 0
    else:
        pdf_title = lines[lines.index("TEST") + 1] if "TEST" in lines else PDF.stem
        section = next((ln for ln in lines if ln.startswith("Section 1")), "")
        m = re.search(r"(\d+) Questions", section)
        EXPECTED_Q = int(m.group(1)) if m else 0
        category = re.search(r"·\s*([^·]+?)\s*·", " ".join(doc[0].get_text("text", clip=pymupdf.Rect(
            0, 795, doc[0].rect.width, doc[0].rect.height)).split()))

    stem = PDF.stem
    out_dir = Path(out_dir) if out_dir else Path.cwd()
    out_dir.mkdir(parents=True, exist_ok=True)
    OUT_DOCX = unique_path(out_dir / f"{stem}_Fully_Editable.docx", overwrite)
    suffix = OUT_DOCX.stem[len(f"{stem}_Fully_Editable"):]  # "" or " (2)" ...
    REVIEW = out_dir / f"conversion_review{suffix}.txt"
    IMG_DIR = out_dir / "extracted_images"

    FORMULA_IMAGES, FIGURE_TABLES, SOLUTIONS, REVIEW_NOTES, REVIEW_GENERAL = {}, {}, {}, {}, []
    UNCLEAR = "[UNCLEAR IN SOURCE]"
    if transcriptions:
        TRANSCRIPTION_FILE = Path(transcriptions)
        if not TRANSCRIPTION_FILE.exists():
            raise ConversionError(f"Transcription module not found: {TRANSCRIPTION_FILE}")
    else:
        TRANSCRIPTION_FILE = HERE / f"transcriptions_{stem}.py"
        if not TRANSCRIPTION_FILE.exists() and stem == "TEST-29":
            TRANSCRIPTION_FILE = HERE / "solution_transcriptions.py"
    TITLE = pdf_title
    if TRANSCRIPTION_FILE.exists():
        try:
            t = load_transcriptions(TRANSCRIPTION_FILE)
        except Exception as exc:
            raise ConversionError(f"Cannot load transcription module {TRANSCRIPTION_FILE.name}: {exc}") from exc
        FORMULA_IMAGES = getattr(t, "FORMULA_IMAGES", {})
        FIGURE_TABLES = getattr(t, "FIGURE_TABLES", {})
        SOLUTIONS = getattr(t, "SOLUTIONS", {})
        REVIEW_NOTES = getattr(t, "REVIEW_NOTES", {})
        REVIEW_GENERAL = getattr(t, "REVIEW_GENERAL", [])
        UNCLEAR = getattr(t, "UNCLEAR", UNCLEAR)
        TITLE = getattr(t, "TITLE", pdf_title)
    else:
        TRANSCRIPTION_FILE = None
        if AI is None:
            review.append(f"No transcription module found (transcriptions_{stem}.py) and --ai not used; "
                          f"every image is marked {UNCLEAR}.")
    SUBTITLE = ((f"Exam Category: {category.group(1)}  |  " if category else "") +
                section.replace(" : ", ": ").replace(" - ", " – "))
    if GENERIC:
        LOG(f"Layout: generic (no questions/answers/solutions footers) - {len(doc)} pages converted page by page")
    else:
        LOG(f"Layout: questions {Q_PAGES[0]}-{Q_PAGES[-1]}"
            + (f", answers {A_PAGES[0]}-{A_PAGES[-1]}" if A_PAGES else "")
            + (f", solutions {S_PAGES[0]}-{S_PAGES[-1]}" if S_PAGES else "")
            + f"; expected questions: {EXPECTED_Q}")
    LOG(f"Transcriptions: {TRANSCRIPTION_FILE.name if TRANSCRIPTION_FILE else 'none'}"
        f"; AI: {AI.model if AI else 'off'}")


def write_review():
    REVIEW.write_text("CONVERSION REVIEW — items for manual check\n\n" + "\n".join(f"- {x}" for x in review) + "\n",
                      encoding="utf-8")


def convert(pdf, out_dir=None, transcriptions=None, ai=False, model=DEFAULT_MODEL, overwrite=False,
            log=print, transcriber=None, keep_images=False):
    """Convert `pdf` to an editable .docx. Returns a dict describing the outputs.

    ai          - transcribe images that have no manual transcription with Claude (needs credentials)
    transcriber - an object with .model and .transcribe(path, context) (tests inject a fake)
    keep_images - place the original drawings/handwriting/scans in the document as pictures instead of
                  typing them out (printed text stays editable; formula images listed in a manual
                  transcription still become Word equations)
    """
    global PDF, AI, LOG, KEEP_IMAGES
    PDF, LOG, KEEP_IMAGES = Path(pdf), log, keep_images
    review.clear()
    arrow_noted.clear()
    used_formulas.clear()
    ai_transcribed.clear()
    if not PDF.exists():
        raise ConversionError(f"Input PDF not found: {PDF}")
    try:
        doc = pymupdf.open(PDF)
    except Exception as exc:
        raise ConversionError(f"Cannot open {PDF.name} as a PDF: {exc}") from exc
    if doc.needs_pass:
        raise ConversionError(f"{PDF.name} is password-protected; decrypt it first.")
    if len(doc) == 0:
        raise ConversionError(f"{PDF.name} has no pages.")
    out = Path(out_dir) if out_dir else Path.cwd()
    AI = transcriber or (ClaudeTranscriber(model=model, cache_dir=out / "ai_transcriptions") if ai else None)
    ai_requested = AI is not None
    configure(doc, out_dir, transcriptions, overwrite)
    IMG_DIR.mkdir(exist_ok=True)
    n_img = sum(len(p.get_images()) for p in doc)
    LOG(f"PDF: {PDF}  pages={len(doc)}  embedded image refs={n_img}  "
        f"text chars={sum(len(p.get_text()) for p in doc)}")
    imgs = Images(doc)
    result = dict(docx=OUT_DOCX, review=REVIEW, images_dir=IMG_DIR, layout="generic" if GENERIC else "test",
                  pages=len(doc), questions=0, answers=0, solutions=0)

    if GENERIC:
        out_doc = build_generic_docx(doc, imgs)
    else:
        qs, correct = parse_questions(doc, imgs)
        ans = parse_answers(doc, imgs)
        sols = parse_solutions(doc, imgs)
        expect = set(range(1, EXPECTED_Q + 1))
        for name, got in (("questions", set(qs)), ("answers", set(ans)), ("solutions", set(sols))):
            if EXPECTED_Q and got != expect:
                review.append(f"{name}: missing {sorted(expect - got)} unexpected {sorted(got - expect)}")
        norm = lambda runs: " ".join("".join(r[1] for r in runs if r[0] == "t").replace(SUP_ON, "")
                                     .replace(SUP_OFF, "").lower().split())
        for n in sorted(set(qs) & set(ans)):
            if any(r[0] != "t" for r in ans[n]):
                continue  # image answers are checked by their transcription
            options = [norm(p["runs"]) for p in qs[n]["paras"] if p["kind"] == "option"]
            if options and norm(ans[n]) not in options:
                review.append(f"Q{n}: answer-key entry '{norm(ans[n])}' does not match any of its options "
                              f"(pages {qs[n]['pages']}) — kept as in the source.")
        if EXPECTED_Q and correct != EXPECTED_Q:
            review.append(f"'Correct: +N' markers found {correct}, expected {EXPECTED_Q}")
        out_doc = build_docx(qs, ans, sols)
        result.update(questions=len(qs), answers=len(ans), solutions=len(sols))

    method = []
    if KEEP_IMAGES:
        method.append("original images placed in the document (no transcription of image content)")
    if ai_requested and AI is None:
        method.append("Claude AI requested but unavailable (see above)")
    if TRANSCRIPTION_FILE:
        method.append(f"manual visual transcription ({TRANSCRIPTION_FILE.name})")
    if AI is not None:
        method.append(f"Claude vision model {AI.model} ({len(ai_transcribed)} images)")
    review.append("Recognition method: " + (" and ".join(method) or "none — images are marked " + UNCLEAR) +
                  (". Original images are included in the Word document." if KEEP_IMAGES else
                   ". No images are included in the Word document; proof-read each transcription against the "
                   "PDF page given in this file."))
    review.extend(REVIEW_GENERAL)
    out_doc.save(OUT_DOCX)
    write_review()
    result.update(review_items=len(review), ai_images=len(ai_transcribed),
                  unclear=sum(1 for r in review if UNCLEAR in r or "NO transcription" in r))
    LOG(f"Saved {OUT_DOCX}")
    LOG(f"Review: {len(review)} items -> {REVIEW}")
    return result


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf", nargs="?", default="TEST-29.pdf")
    ap.add_argument("--out-dir", help="folder for the .docx, conversion_review.txt and extracted_images/ "
                                      "(default: current folder)")
    ap.add_argument("--transcriptions", help="manual transcription module (.py) to use")
    ap.add_argument("--ai", action="store_true",
                    help="transcribe images without a manual transcription using Claude (needs ANTHROPIC_API_KEY "
                         "or an `ant auth login` profile)")
    ap.add_argument("--model", default=DEFAULT_MODEL, help=f"Claude model for --ai (default {DEFAULT_MODEL})")
    ap.add_argument("--overwrite", action="store_true",
                    help="replace an existing output .docx instead of writing 'name (2).docx'")
    ap.add_argument("--keep-images", action="store_true",
                    help="place the original images (drawings, handwriting, scanned tables) in the document "
                         "instead of typing them out")
    args = ap.parse_args()
    try:
        r = convert(args.pdf, args.out_dir, args.transcriptions, args.ai, args.model, args.overwrite,
                    keep_images=args.keep_images)
    except ConversionError as exc:
        sys.exit(f"Error: {exc}")
    if r["layout"] == "test":
        print(f"Questions: {r['questions']}  Answers: {r['answers']}  Solutions: {r['solutions']}  "
              f"AI-transcribed images: {r['ai_images']}")


if __name__ == "__main__":
    main()
