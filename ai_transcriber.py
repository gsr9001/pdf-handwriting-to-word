"""AI handwriting / formula recognition for convert_pdf_to_word.py, using Claude's vision API.

An image (handwritten solution, sketch, scanned table, printed formula) is sent to Claude, which
returns the content as the same block markup the manual transcription modules use:

  {"kind": "eq",      "text": "<equation markup>", "rows": []}
  {"kind": "p",       "text": "text with $inline$ maths", "rows": []}
  {"kind": "table",   "text": "", "rows": [["header", ...], ["cell", ...]]}
  {"kind": "diagram", "text": "", "rows": [["label", "→", "label"], ...]}
  {"kind": "note",    "text": "transcriber's note", "rows": []}

plus a list of "uncertain" items for the review report. Responses are constrained with structured
outputs (JSON schema), validated here, and cached as JSON next to the output so re-runs do not
re-bill and a person can correct a transcription by editing the cached file.

Credentials come from the environment (ANTHROPIC_API_KEY, or an `ant auth login` profile);
nothing is hard-coded.
"""
import base64
import json
import mimetypes
from pathlib import Path

DEFAULT_MODEL = "claude-opus-5-5"
BLOCK_KINDS = ("eq", "p", "table", "diagram", "note")

SYSTEM_PROMPT = """You transcribe images from engineering examination PDFs (handwritten solutions, \
calculations, sketches, scanned tables, printed formulas) into editable structured content.

Rules:
- Transcribe exactly what is written, in the order written. Do not correct spelling, arithmetic, \
formulas or answers, and do not add steps, explanations or answers that are not in the image.
- Where something cannot be read confidently, write [UNCLEAR IN SOURCE] in its place and add a short \
entry to "uncertain" saying what and where (e.g. "line 3: value after 'e =' partly cut off").
- One block per line or step of working. Use kind "eq" for a line that is purely an equation, "p" \
for text (wrap any maths inside it in $...$), "table" for a table (first row = header), "diagram" \
for a simple sketch laid out as a small grid of labels and arrow characters (→ ← ↑ ↓ ↗ ↘ ↙ ↖), and \
"note" for describing anything that cannot be reproduced (a drawing, crossed-out text, a circled \
sign). Leave "rows" empty except for table and diagram; leave "text" empty for table and diagram.
- Equation markup: \\frac{a}{b} for fractions, x_{sub} and x^{sup} for sub/superscripts, \
\\box{...} for boxed expressions, \\text{words} for upright words or units inside an equation. \
Write Greek letters and symbols directly (γ ρ η Δ ≤ ≥ × ∝ ∴ ⇒ °).
- Do not wrap the JSON in code fences."""

RESPONSE_SCHEMA = {
    "type": "object",
    "properties": {
        "blocks": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "kind": {"type": "string", "enum": list(BLOCK_KINDS)},
                    "text": {"type": "string"},
                    "rows": {"type": "array", "items": {"type": "array", "items": {"type": "string"}}},
                },
                "required": ["kind", "text", "rows"],
                "additionalProperties": False,
            },
        },
        "uncertain": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["blocks", "uncertain"],
    "additionalProperties": False,
}


class TranscriptionError(RuntimeError):
    """The image could not be transcribed (API error, refusal, invalid response).

    `fatal` is True when no further image can succeed either (missing or rejected credentials)."""

    def __init__(self, message, fatal=False):
        super().__init__(message)
        self.fatal = fatal


NO_CREDENTIALS = ("no valid Claude credentials - set ANTHROPIC_API_KEY or run `ant auth login`")


def validate_response(data):
    """Check a parsed response and convert it to (blocks, uncertain) in the converter's tuple form."""
    if not isinstance(data, dict) or not isinstance(data.get("blocks"), list):
        raise TranscriptionError("response has no 'blocks' list")
    blocks = []
    for i, b in enumerate(data["blocks"]):
        kind = b.get("kind") if isinstance(b, dict) else None
        if kind not in BLOCK_KINDS:
            raise TranscriptionError(f"block {i} has invalid kind {kind!r}")
        if kind in ("table", "diagram"):
            rows = b.get("rows") or []
            if not rows or not all(isinstance(r, list) and r for r in rows):
                raise TranscriptionError(f"block {i} ({kind}) has no rows")
            width = max(len(r) for r in rows)
            rows = [[str(c) for c in r] + [""] * (width - len(r)) for r in rows]  # pad ragged rows
            blocks.append((kind, rows))
        else:
            text = str(b.get("text", "")).strip()
            if text:
                blocks.append((kind, text))
    if not blocks:
        raise TranscriptionError("response contains no content")
    uncertain = [str(u) for u in data.get("uncertain", []) if str(u).strip()]
    return blocks, uncertain


class ClaudeTranscriber:
    """Transcribes images with Claude. `client` may be injected (tests pass a fake)."""

    def __init__(self, model=DEFAULT_MODEL, client=None, cache_dir=None, effort="high"):
        self.model = model
        self.effort = effort
        self.cache_dir = Path(cache_dir) if cache_dir else None
        self._client = client
        self.calls = 0

    @property
    def client(self):
        if self._client is None:
            try:
                import anthropic
            except ImportError as exc:
                raise TranscriptionError("the 'anthropic' package is not installed (pip install anthropic)") from exc
            self._client = anthropic.Anthropic()
        return self._client

    def _cache_path(self, image_path):
        return self.cache_dir / f"{Path(image_path).stem}.json" if self.cache_dir else None

    def transcribe(self, image_path, context=""):
        """Return (blocks, uncertain, from_cache)."""
        cache = self._cache_path(image_path)
        if cache and cache.exists():
            blocks, uncertain = validate_response(json.loads(cache.read_text(encoding="utf-8")))
            return blocks, uncertain, True
        data = self._request(Path(image_path), context)
        blocks, uncertain = validate_response(data)
        if cache:
            cache.parent.mkdir(parents=True, exist_ok=True)
            cache.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return blocks, uncertain, False

    def _request(self, image_path, context):
        media_type = mimetypes.guess_type(image_path.name)[0] or "image/png"
        image_b64 = base64.standard_b64encode(image_path.read_bytes()).decode("ascii")
        prompt = "Transcribe this image."
        if context:
            prompt += f" Context: {context}."
        try:
            import anthropic
            errors = (anthropic.APIError,)
            auth_errors = (anthropic.AuthenticationError, anthropic.PermissionDeniedError)
        except ImportError:
            errors = auth_errors = ()
        try:
            response = self.client.beta.messages.create(
                model=self.model,
                max_tokens=16000,
                betas=["server-side-fallback-2026-07-01"],
                fallbacks="default",
                system=SYSTEM_PROMPT,
                output_config={"effort": self.effort, "format": {"type": "json_schema", "schema": RESPONSE_SCHEMA}},
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": image_b64}},
                    {"type": "text", "text": prompt},
                ]}],
            )
        except auth_errors as exc:
            raise TranscriptionError(f"{NO_CREDENTIALS} ({exc})", fatal=True) from exc
        except errors as exc:
            raise TranscriptionError(f"Claude API error for {image_path.name}: {exc}") from exc
        except TypeError as exc:  # raised by the SDK when no credential source resolves
            if "authentication" in str(exc).lower():
                raise TranscriptionError(NO_CREDENTIALS, fatal=True) from exc
            raise
        self.calls += 1
        if response.stop_reason == "refusal":
            raise TranscriptionError(f"Claude declined to transcribe {image_path.name}")
        if response.stop_reason == "max_tokens":
            raise TranscriptionError(f"response for {image_path.name} was cut off (max_tokens)")
        text = next((b.text for b in response.content if b.type == "text"), "")
        try:
            return json.loads(text)
        except json.JSONDecodeError as exc:
            raise TranscriptionError(f"invalid JSON returned for {image_path.name}") from exc
