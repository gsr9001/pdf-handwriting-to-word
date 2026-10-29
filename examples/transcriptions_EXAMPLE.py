"""Example manual transcription module (synthetic content).

Copy this file next to convert_pdf_to_word.py as transcriptions_<PDF file stem>.py (for example
transcriptions_TEST-41.py for TEST-41.pdf), or pass it with --transcriptions. Every key is
optional; anything not transcribed here is sent to Claude when --ai is used, otherwise it is
marked [UNCLEAR IN SOURCE] and listed in the review report.

Block markup
  ("eq", m)            display equation (Word equation)
  ("p", s)             paragraph; segments wrapped in $...$ are inline equations
  ("table", rows)      Word table; first row is the header; a row whose other cells are "" spans
                       the full width
  ("diagram", rows)    borderless centred grid of labels and arrows; a cell "<<" merges into the
                       cell on its left
  ("image", path[, w])  a redrawn figure picture (width in inches, max 6)
  ("note", s)          transcriber's note (grey italics; not part of the source)
Equation syntax: \\frac{a}{b}, x_{sub}, x^{sup}, \\box{...}, \\text{upright words}
"""

UNCLEAR = "[UNCLEAR IN SOURCE]"

# Optional document title (defaults to the title printed in the PDF).
TITLE = "Test-99 (Synthetic Sample)"

# Printed formula images inside question/answer text, keyed by extracted image file name
# (see the extracted_images/ folder after a first run).
FORMULA_IMAGES = {
    "question_q02_formula1.png": r"x=\frac{10}{2}",
}

# Image tables/figures inside questions, keyed by extracted image file name -> table rows.
FIGURE_TABLES = {
    "question_q05_table.png": [["Column A", "Column B"], ["1. first item", "a. its match"]],
}

# Handwritten / image solutions, keyed by question number.
SOLUTIONS = {
    1: [
        ("eq", r"e=\frac{V_{v}}{V_{s}}=\frac{4}{8}"),
        ("p", "Void ratio is $e$, porosity is $η=\\frac{e}{1+e}$"),
        ("table", [["Property", "Valid range"], ["$e$", "$e>0$"]]),
        ("diagram", [["Step A", "→", "Step B"]]),
        ("note", "A word in the margin is crossed out in the source."),
    ],
}

# Per-solution notes copied into the review report (with the PDF page number).
REVIEW_NOTES = {
    1: "value in line 2 is smudged - read as 0.5",
}

# General notes copied into the review report.
REVIEW_GENERAL = [
    "Printed text kept verbatim.",
]
