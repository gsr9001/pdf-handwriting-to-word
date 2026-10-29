"""ClaudeTranscriber with a mocked Anthropic client - no network calls, no API key needed."""
import json
from types import SimpleNamespace

import pytest

from ai_transcriber import RESPONSE_SCHEMA, ClaudeTranscriber, TranscriptionError, validate_response
from conftest import png_bytes

GOOD = {
    "blocks": [
        {"kind": "eq", "text": "e=\\frac{V_{v}}{V_{s}}", "rows": []},
        {"kind": "p", "text": "Void ratio [UNCLEAR IN SOURCE]", "rows": []},
        {"kind": "table", "text": "", "rows": [["Property", "Range"], ["e"]]},
    ],
    "uncertain": ["line 2: value smudged"],
}


class FakeMessages:
    def __init__(self, text, stop_reason="end_turn"):
        self.text, self.stop_reason, self.requests = text, stop_reason, []

    def create(self, **kwargs):
        self.requests.append(kwargs)
        return SimpleNamespace(stop_reason=self.stop_reason,
                               content=[SimpleNamespace(type="text", text=self.text)])


def fake_client(text, stop_reason="end_turn"):
    messages = FakeMessages(text, stop_reason)
    return SimpleNamespace(beta=SimpleNamespace(messages=messages)), messages


@pytest.fixture
def image(tmp_path):
    p = tmp_path / "solution_q01_1.png"
    p.write_bytes(png_bytes("e = Vv / Vs"))
    return p


def test_request_shape_and_parsing(image, tmp_path):
    client, messages = fake_client(json.dumps(GOOD))
    t = ClaudeTranscriber(client=client, cache_dir=tmp_path / "cache")
    blocks, uncertain, cached = t.transcribe(image, "Solution 1 (PDF page 21)")
    assert not cached
    assert blocks[0] == ("eq", "e=\\frac{V_{v}}{V_{s}}")
    assert blocks[2] == ("table", [["Property", "Range"], ["e", ""]])  # ragged row padded
    assert uncertain == ["line 2: value smudged"]
    req = messages.requests[0]
    assert req["model"] == "claude-opus-5-5"
    assert req["output_config"]["format"] == {"type": "json_schema", "schema": RESPONSE_SCHEMA}
    assert req["fallbacks"] == "default"
    content = req["messages"][0]["content"]
    assert content[0]["type"] == "image" and content[0]["source"]["media_type"] == "image/png"
    assert "Solution 1" in content[1]["text"]


def test_cache_is_written_and_reused(image, tmp_path):
    client, messages = fake_client(json.dumps(GOOD))
    t = ClaudeTranscriber(client=client, cache_dir=tmp_path / "cache")
    t.transcribe(image)
    assert (tmp_path / "cache" / "solution_q01_1.json").exists()
    _, _, cached = t.transcribe(image)
    assert cached and len(messages.requests) == 1


@pytest.mark.parametrize("text,stop,msg", [
    ("{}", "refusal", "declined"),
    ('{"blocks": [', "max_tokens", "cut off"),
    ("not json", "end_turn", "invalid JSON"),
])
def test_bad_responses_raise(image, text, stop, msg):
    client, _ = fake_client(text, stop)
    with pytest.raises(TranscriptionError, match=msg):
        ClaudeTranscriber(client=client).transcribe(image)


def test_api_errors_are_wrapped(image):
    import anthropic
    import httpx2

    class Failing:
        def create(self, **_):
            raise anthropic.APIConnectionError(request=httpx2.Request("POST", "https://api.anthropic.com"))

    client = SimpleNamespace(beta=SimpleNamespace(messages=Failing()))
    with pytest.raises(TranscriptionError, match="API error"):
        ClaudeTranscriber(client=client).transcribe(image)


@pytest.mark.parametrize("data", [
    {"uncertain": []},
    {"blocks": [{"kind": "image", "text": "x", "rows": []}], "uncertain": []},
    {"blocks": [{"kind": "table", "text": "", "rows": []}], "uncertain": []},
    {"blocks": [{"kind": "p", "text": "  ", "rows": []}], "uncertain": []},
])
def test_validate_rejects_bad_content(data):
    with pytest.raises(TranscriptionError):
        validate_response(data)


def test_authentication_errors_are_fatal(image):
    class NoAuth:
        def create(self, **_):
            raise TypeError("Could not resolve authentication method.")

    client = SimpleNamespace(beta=SimpleNamespace(messages=NoAuth()))
    with pytest.raises(TranscriptionError) as info:
        ClaudeTranscriber(client=client).transcribe(image)
    assert info.value.fatal
