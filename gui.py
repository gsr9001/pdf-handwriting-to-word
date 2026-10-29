"""Desktop window for convert_pdf_to_word.py (Tkinter - no extra dependencies).

Run:  python gui.py
"""
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import convert_pdf_to_word as conv
from ai_transcriber import DEFAULT_MODEL


class ConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Handwriting → Editable Word")
        self.minsize(640, 460)
        self.messages = queue.Queue()
        self.worker = None
        self.last_result = None

        self.pdf = tk.StringVar()
        self.out_dir = tk.StringVar()
        self.transcriptions = tk.StringVar()
        self.use_ai = tk.BooleanVar(value=False)
        self.model = tk.StringVar(value=DEFAULT_MODEL)
        self.overwrite = tk.BooleanVar(value=False)
        self.keep_images = tk.BooleanVar(value=False)

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)
        self._path_row(frm, 0, "PDF file", self.pdf, self._pick_pdf)
        self._path_row(frm, 1, "Output folder", self.out_dir, self._pick_out)
        self._path_row(frm, 2, "Transcriptions (optional)", self.transcriptions, self._pick_trans)

        ai = ttk.Frame(frm)
        ai.grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Checkbutton(ai, text="Use Claude AI for images without a manual transcription",
                        variable=self.use_ai).pack(side="left")
        ttk.Label(ai, text="  Model:").pack(side="left")
        ttk.Entry(ai, textvariable=self.model, width=20).pack(side="left")
        opts = ttk.Frame(frm)
        opts.grid(row=4, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(opts, text="Overwrite an existing output .docx", variable=self.overwrite).pack(side="left")
        ttk.Checkbutton(opts, text="  Keep original images (drawings, handwriting) instead of typing them",
                        variable=self.keep_images).pack(side="left")

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=3, sticky="w", pady=8)
        self.convert_btn = ttk.Button(btns, text="Convert", command=self._start)
        self.convert_btn.pack(side="left")
        self.open_doc_btn = ttk.Button(btns, text="Open document", command=lambda: self._open("docx"),
                                       state="disabled")
        self.open_doc_btn.pack(side="left", padx=6)
        self.open_rev_btn = ttk.Button(btns, text="Open review report", command=lambda: self._open("review"),
                                       state="disabled")
        self.open_rev_btn.pack(side="left")
        self.progress = ttk.Progressbar(frm, mode="indeterminate")
        self.progress.grid(row=6, column=0, columnspan=3, sticky="ew")

        self.log = tk.Text(frm, height=14, wrap="word", state="disabled")
        self.log.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        frm.rowconfigure(7, weight=1)
        self.after(100, self._drain)

    def _path_row(self, parent, row, label, var, command):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=2)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=6)
        ttk.Button(parent, text="Browse…", command=command).grid(row=row, column=2)

    def _pick_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.pdf.set(path)
            if not self.out_dir.get():
                self.out_dir.set(str(Path(path).parent))

    def _pick_out(self):
        path = filedialog.askdirectory()
        if path:
            self.out_dir.set(path)

    def _pick_trans(self):
        path = filedialog.askopenfilename(filetypes=[("Python transcription module", "*.py")])
        if path:
            self.transcriptions.set(path)

    def _write(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _start(self):
        if not self.pdf.get():
            messagebox.showwarning("No PDF", "Choose a PDF file first.")
            return
        if self.use_ai.get() and not (os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")):
            if not messagebox.askyesno("Claude credentials",
                                       "ANTHROPIC_API_KEY is not set. Continue anyway (works if you have run "
                                       "`ant auth login`)?"):
                return
        self.convert_btn.configure(state="disabled")
        self.open_doc_btn.configure(state="disabled")
        self.open_rev_btn.configure(state="disabled")
        self.progress.start(12)
        self._write(f"Converting {self.pdf.get()} …")
        args = dict(pdf=self.pdf.get(), out_dir=self.out_dir.get() or None,
                    transcriptions=self.transcriptions.get() or None, ai=self.use_ai.get(),
                    model=self.model.get().strip() or DEFAULT_MODEL, overwrite=self.overwrite.get(),
                    keep_images=self.keep_images.get(), log=lambda m: self.messages.put(("log", m)))
        self.worker = threading.Thread(target=self._run, args=(args,), daemon=True)
        self.worker.start()

    def _run(self, args):
        try:
            self.messages.put(("done", conv.convert(**args)))
        except conv.ConversionError as exc:
            self.messages.put(("error", str(exc)))
        except Exception as exc:  # show unexpected failures instead of dying silently
            self.messages.put(("error", f"{type(exc).__name__}: {exc}"))

    def _drain(self):
        try:
            while True:
                kind, payload = self.messages.get_nowait()
                if kind == "log":
                    self._write(payload)
                    continue
                self.progress.stop()
                self.convert_btn.configure(state="normal")
                if kind == "error":
                    self._write(f"ERROR: {payload}")
                    messagebox.showerror("Conversion failed", payload)
                else:
                    self.last_result = payload
                    self.open_doc_btn.configure(state="normal")
                    self.open_rev_btn.configure(state="normal")
                    self._write(f"Done. {payload['review_items']} review items "
                                f"({payload['unclear']} unreadable/untranscribed).")
        except queue.Empty:
            pass
        self.after(100, self._drain)

    def _open(self, key):
        path = self.last_result and self.last_result.get(key)
        if not path:
            return
        if sys.platform.startswith("win"):
            os.startfile(path)
        else:
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)])


if __name__ == "__main__":
    ConverterApp().mainloop()
