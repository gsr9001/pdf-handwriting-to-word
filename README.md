# PDF Handwriting → Editable Word Converter

Converts exam PDFs that mix printed text with handwritten solutions, formulas, sketches and
scanned tables into a **fully editable, professionally formatted Word document**. Printed text is
copied from the PDF; image content is typed as Word text, Word equations, Word tables and simple
text/arrow diagrams. **No original images or scanned pages are placed in the output.** Anything
that cannot be read is marked `[UNCLEAR IN SOURCE]` and listed, with its page, in a review report.

## Features

- **Printed text extraction** (PyMuPDF), keeping superscripts (m³, ⁰C) and symbol-font arrows.
- **Exam layout detection** for question / answer-key / solution PDFs (Classplus-style footers):
  question numbers (including numbers printed beside a wrapped stem), options (radio buttons),
  answer key, and solutions are rebuilt in order. Missing question numbers, missing stems, option
  counts that differ from four and answer-key entries that match no option are reported.
- **Generic mode** for any other PDF: each page is converted in order; printed text is kept and
  embedded images or pen-stroke pages are transcribed.
- **Handwriting and formula recognition**, in this order of precedence:
  1. **Manual transcriptions** — a Python module of typed blocks (exact, reviewable, version-free);
  2. **Claude vision (`--ai`)** — each image is sent to Claude, which returns structured blocks
     (JSON-schema constrained) plus a list of uncertain readings. Results are cached as JSON in
     `<out-dir>/ai_transcriptions/`, so re-runs cost nothing and corrections can be made by
     editing the cached file;
  3. otherwise the item is marked `[UNCLEAR IN SOURCE]`.
- **Word output**: A4, 1" margins, Calibri 11 pt, numbered questions with indented (A)–(D)
  options, answer-key table, native Word equations (fractions, sub/superscripts, boxes),
  editable tables, header with the test title and `Page X of Y` footer.
- **Review report** (`conversion_review.txt`): unreadable handwriting, AI-flagged uncertainties,
  source inconsistencies, missing content — each with question/solution number and PDF page.
- **Desktop UI** (Tkinter) and **command line**.
- **Never overwrites**: an existing output is kept and the new one is written as `name (2).docx`
  (use `--overwrite` to replace).

## Installation

Python 3.10+.

```bash
pip install -r requirements.txt          # runtime
pip install -r requirements-dev.txt      # + pytest
```

For `--ai`, provide Claude credentials through the environment — never in code:

```bash
export ANTHROPIC_API_KEY=...             # PowerShell: $env:ANTHROPIC_API_KEY = "..."
# or: ant auth login                      (the SDK picks up the stored profile)
```

## Usage

### Desktop app

```bash
python gui.py
```

Choose the PDF, an output folder, optionally a manual transcription module, tick
**Use Claude AI** if wanted, and press **Convert**. Buttons open the document and the review
report when it finishes.

### Command line

```bash
# manual transcriptions only (transcriptions_<stem>.py is picked up automatically)
python convert_pdf_to_word.py TEST-41.pdf --out-dir out/

# let Claude transcribe every image that has no manual transcription
python convert_pdf_to_word.py TEST-41.pdf --out-dir out/ --ai

# options
#   --transcriptions FILE.py   use a specific manual transcription module
#   --model MODEL              Claude model for --ai (default claude-opus-5-5)
#   --overwrite                replace an existing output .docx
#   --keep-images              put the original drawings/handwriting/scans in the document as
#                              pictures instead of typing them out
```

Outputs in the output folder:

| File | Contents |
|---|---|
| `<stem>_Fully_Editable.docx` | the editable document (no images, unless `--keep-images` is used) |
| `conversion_review.txt` | items to check by hand |
| `extracted_images/` | the PDF's images, extracted for reference/transcription (not in the .docx unless `--keep-images`) |
| `ai_transcriptions/` | cached Claude transcriptions (with `--ai`), editable JSON |

From Python:

```python
from convert_pdf_to_word import convert
result = convert("TEST-41.pdf", out_dir="out", ai=True)
print(result["docx"], result["review"], result["unclear"])
```

## Manual transcription modules

Exact, reviewed transcriptions live in `transcriptions_<PDF stem>.py` next to the script (or are
passed with `--transcriptions`). See [`examples/transcriptions_EXAMPLE.py`](examples/transcriptions_EXAMPLE.py)
for every supported key and the block markup:

```python
SOLUTIONS = {
    12: [
        ("eq", r"e=\frac{V_{v}}{V_{s}}=\frac{4}{8}"),
        ("p", "Void ratio is $e$ and porosity is $η=\\frac{e}{1+e}$"),
        ("table", [["Property", "Valid range"], ["$e$", "$e>0$"]]),
        ("diagram", [["Step A", "→", "Step B"]]),
        ("note", "A word in the margin is crossed out in the source."),
    ],
}
```

Equation markup: `\frac{a}{b}`, `x_{sub}`, `x^{sup}`, `\box{...}`, `\text{upright words}`;
Greek letters and symbols are typed directly. Image file names for `FORMULA_IMAGES` /
`FIGURE_TABLES` are the names written to `extracted_images/` on a first run.

Transcription modules of real papers contain the papers' content and are **git-ignored**.

## How accurate is it?

- Printed text is exact (it is the PDF's own text layer).
- Manual transcriptions are as accurate as the person who typed them.
- AI transcriptions are a draft: Claude is instructed to transcribe literally, never correct
  arithmetic, and flag what it cannot read, but every AI-transcribed item is listed in the
  review report and must be proof-read against the PDF.
- Sketches are reproduced as text/arrow grids or described in a note — not redrawn.

## Tests

```bash
python -m pytest
```

The suite builds synthetic PDFs on the fly (no private papers needed) and uses a mocked
transcriber / mocked Anthropic client, so it makes **no API calls** and needs no key. It covers
the equation markup, layout detection, the AI request/response handling (schema, caching,
refusals, invalid JSON, API errors), end-to-end conversion in both layouts, no-image output,
no-overwrite behaviour, error handling and a GUI smoke test.

## Project layout

| Path | Purpose |
|---|---|
| `convert_pdf_to_word.py` | conversion pipeline + CLI (`convert()` is the API) |
| `ai_transcriber.py` | Claude vision transcriber (structured output, validation, cache) |
| `gui.py` | Tkinter desktop app |
| `examples/transcriptions_EXAMPLE.py` | documented manual-transcription template |
| `tests/` | pytest suite with synthetic PDFs and mocks |

## Privacy

Source PDFs, generated documents, review reports, extracted images, AI caches and transcription
modules of real papers are excluded by `.gitignore`. With `--ai`, images are sent to the
Anthropic API; do not use `--ai` for material you are not permitted to share with it.

## License

License: MIT (see [LICENSE](LICENSE)).
