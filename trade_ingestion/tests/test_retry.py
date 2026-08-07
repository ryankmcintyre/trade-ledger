from __future__ import annotations

from trade_ingestion.models import ResolutionFailure
from trade_ingestion.retry import resolve_with_retry


def test_resolve_with_retry_succeeds_on_first_try() -> None:
    def try_resolve(value: str) -> tuple[str | None, str | None]:
        return (value.upper(), None) if value == "aapl" else (None, "not found")

    result, failure = resolve_with_retry("aapl", "record-1", try_resolve, prompt=None)

    assert result == "AAPL"
    assert failure is None


def test_resolve_with_retry_succeeds_after_prompt_retry() -> None:
    def try_resolve(value: str) -> tuple[str | None, str | None]:
        return (value, None) if value == "HON" else (None, f"unsupported: {value}")

    def prompt(input_value: str, context_label: str) -> str | None:
        assert input_value == "HON2"
        assert context_label == "record-1"
        return "HON"

    result, failure = resolve_with_retry("HON2", "record-1", try_resolve, prompt)

    assert result == "HON"
    assert failure is None


def test_resolve_with_retry_records_failure_when_no_prompt_available() -> None:
    def try_resolve(value: str) -> tuple[str | None, str | None]:
        return None, "unsupported format"

    result, failure = resolve_with_retry("HON2", "record-1", try_resolve, prompt=None)

    assert result is None
    assert failure == ResolutionFailure(
        context_label="record-1",
        input_value="HON2",
        attempted_value="HON2",
        error="unsupported format",
    )


def test_resolve_with_retry_records_failure_when_prompt_declines() -> None:
    def try_resolve(value: str) -> tuple[str | None, str | None]:
        return None, "unsupported format"

    def prompt(input_value: str, context_label: str) -> str | None:
        return None

    result, failure = resolve_with_retry("HON2", "record-1", try_resolve, prompt)

    assert result is None
    assert failure is not None
    assert failure.attempted_value == "HON2"
    assert failure.error == "unsupported format"


def test_resolve_with_retry_records_failure_when_retry_still_fails() -> None:
    def try_resolve(value: str) -> tuple[str | None, str | None]:
        return None, f"still unsupported: {value}"

    def prompt(input_value: str, context_label: str) -> str | None:
        return "HONX"

    result, failure = resolve_with_retry("HON2", "record-1", try_resolve, prompt)

    assert result is None
    assert failure is not None
    assert failure.attempted_value == "HONX"
    assert failure.error == "still unsupported: HONX"


def test_resolve_with_retry_treats_blank_prompt_response_as_no_replacement() -> None:
    def try_resolve(value: str) -> tuple[str | None, str | None]:
        return None, "unsupported format"

    def prompt(input_value: str, context_label: str) -> str | None:
        return "   "

    result, failure = resolve_with_retry("HON2", "record-1", try_resolve, prompt)

    assert result is None
    assert failure is not None
    assert failure.attempted_value == "HON2"


def test_resolve_with_retry_handles_prompt_exception() -> None:
    def try_resolve(value: str) -> tuple[str | None, str | None]:
        return None, "unsupported format"

    def prompt(input_value: str, context_label: str) -> str | None:
        raise RuntimeError("boom")

    result, failure = resolve_with_retry("HON2", "record-1", try_resolve, prompt)

    assert result is None
    assert failure is not None
    assert "Prompt failed" in failure.error
