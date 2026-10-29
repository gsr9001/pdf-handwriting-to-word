"""Smoke test of the Tkinter window (skipped when no display is available)."""
import pytest

tk = pytest.importorskip("tkinter")


@pytest.fixture
def app():
    gui = pytest.importorskip("gui")
    try:
        window = gui.ConverterApp()
    except tk.TclError as exc:
        pytest.skip(f"no display: {exc}")
    window.withdraw()
    yield window
    window.destroy()


def test_conversion_runs_from_gui(app, test_pdf, tmp_path):
    app.pdf.set(str(test_pdf))
    app.out_dir.set(str(tmp_path / "gui-out"))
    app._run(dict(pdf=app.pdf.get(), out_dir=app.out_dir.get(), transcriptions=None, ai=False,
                  model="m", overwrite=False, log=lambda m: app.messages.put(("log", m))))
    app._drain()
    assert app.last_result and app.last_result["docx"].exists()
    assert str(app.open_doc_btn["state"]) == "normal"
    assert "Done." in app.log.get("1.0", "end")


def test_errors_are_shown_not_raised(app, tmp_path, monkeypatch):
    shown = []
    monkeypatch.setattr("gui.messagebox.showerror", lambda *a: shown.append(a))
    app._run(dict(pdf=str(tmp_path / "missing.pdf"), out_dir=None, transcriptions=None, ai=False,
                  model="m", overwrite=False, log=lambda m: None))
    app._drain()
    assert shown and "not found" in shown[0][1]
