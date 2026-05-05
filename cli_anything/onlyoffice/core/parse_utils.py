"""Small CLI parsing helpers shared by modality handlers."""

from __future__ import annotations

import math
from typing import Any, Callable


def parse_int(value: str, option: str) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        raise ValueError(f"Invalid {option}: expected integer, got {value!r}") from None


def parse_float(value: str, option: str) -> float:
    try:
        parsed = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"Invalid {option}: expected number, got {value!r}") from None
    if not math.isfinite(parsed):
        raise ValueError(f"Invalid {option}: expected finite number, got {value!r}") from None
    return parsed


def print_usage_error(
    print_result: Callable[[dict, bool], None],
    json_output: bool,
    message: str,
    *,
    usage: str | None = None,
    code: str = "usage_error",
) -> None:
    payload: dict[str, Any] = {
        "success": False,
        "error": message,
        "error_code": code,
    }
    if usage:
        payload["usage"] = usage
    print_result(payload, json_output)
