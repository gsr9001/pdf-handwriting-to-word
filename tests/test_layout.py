"""Layout helpers: question-number detection, row grouping, output naming."""
import convert_pdf_to_word as conv


def el(text, x0, y0, x1=None, y1=None):
    return dict(kind="t", x0=x0, y0=y0, x1=x1 or x0 + 100, y1=y1 or y0 + 16, text=text)


def test_split_number_separate_and_fused():
    n, row = conv.split_number(dict(els=[el("7", 36, 100), el("Sample stem", 54, 100)], x0=36), 50)
    assert n == 7 and row["els"][0]["text"] == "Sample stem"
    n, row = conv.split_number(dict(els=[el("12 Sample question text", 36, 100)], x0=36), 50)
    assert n == 12 and row["els"][0]["text"] == "Sample question text" and row["x0"] == 36
    n, _ = conv.split_number(dict(els=[el("3", 63, 100)], x0=63), 50)  # an option "3", not a number
    assert n is None


def test_lift_number_to_first_line_of_wrapped_stem():
    number = el("5", 36, 148, 42, 162)
    lines = [el("First line of a wrapped stem", 48, 130, 485, 146), el("_____", 48, 146, 184, 162)]
    out = conv.lift_question_numbers([*lines, number])
    assert out[0]["text"] in ("5", "First line of a wrapped stem")
    assert number["y0"] < 140  # moved up to the first stem line
    rows = conv.group_rows(out)
    assert rows[0]["els"][0]["text"] == "5"


def test_group_rows_joins_inline_image_with_text():
    img = dict(kind="i", x0=243, y0=95, x1=251, y1=110, xref=1, digest=b"")
    rows = conv.group_rows([el("text before image", 60, 92, 243, 108), img, el("text after image", 251, 92, 500, 108)])
    assert len(rows) == 1 and [e["kind"] for e in rows[0]["els"]] == ["t", "i", "t"]


def test_unique_path(tmp_path):
    p = tmp_path / "out.docx"
    assert conv.unique_path(p, False) == p
    p.write_text("x")
    assert conv.unique_path(p, False).name == "out (2).docx"
    (tmp_path / "out (2).docx").write_text("x")
    assert conv.unique_path(p, False).name == "out (3).docx"
    assert conv.unique_path(p, True) == p
