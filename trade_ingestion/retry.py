"""Shared "try, prompt for a replacement, retry once" resolution algorithm.

This module has no dependency on Excel or CSV parsing — it exists so that any
adapter/writer code that needs to recover from a failed resolution (an
unsupported ticker, an unparseable symbol, etc.) by prompting the operator for
a replacement value and retrying once, shares one implementation instead of
each caller hand-rolling its own retry loop.
"""

from __future__ import annotations

from typing import Callable, TypeVar

from trade_ingestion.models import ResolutionFailure

T = TypeVar("T")

# A resolver attempts to turn an input value into a result. It returns
# (result, None) on success, or (None, error_message) on failure.
Resolver = Callable[[str], "tuple[T | None, str | None]"]
# A prompt asks the operator for a replacement value given the original input
# and a human-readable label identifying the record being resolved. Returning
# None (or an empty/whitespace string) means "no replacement, give up".
Prompt = Callable[[str, str], "str | None"]


def resolve_with_retry(
    input_value: str,
    context_label: str,
    try_resolve: Resolver[T],
    prompt: Prompt | None = None,
) -> tuple[T | None, ResolutionFailure[T] | None]:
    """Attempt ``try_resolve(input_value)``; on failure, optionally prompt once
    for a replacement value and retry with it.

    Returns ``(result, None)`` on success (first try or after a successful
    retry), or ``(None, ResolutionFailure)`` if the initial attempt failed and
    either no prompt was available, the prompt yielded no usable replacement,
    or the retry with the replacement also failed.
    """
    result, error = try_resolve(input_value)
    if result is not None:
        return result, None

    if prompt is None:
        return None, ResolutionFailure(context_label, input_value, input_value, error or "Resolution failed")

    try:
        replacement = prompt(input_value, context_label)
    except Exception as exc:  # noqa: BLE001 - prompt failures are best-effort and must not abort ingestion
        return None, ResolutionFailure(context_label, input_value, input_value, f"Prompt failed: {exc}")

    replacement_text = (replacement or "").strip()
    if not replacement_text:
        return None, ResolutionFailure(context_label, input_value, input_value, error or "Resolution failed")

    retry_result, retry_error = try_resolve(replacement_text)
    if retry_result is not None:
        return retry_result, None

    return None, ResolutionFailure(
        context_label,
        input_value,
        replacement_text,
        retry_error or error or "Resolution failed",
    )
