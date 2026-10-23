"""Clean text cells — remove symbols while keeping letters, numbers, and decimals."""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Any

_MULTI_SPACE = re.compile(r"\s+")


def remove_symbols(text: str) -> str:
  """Drop special characters; keep letters, digits, spaces, and decimal points."""
  if not text:
    return ""
  cleaned = "".join(
    ch for ch in text
    if ch.isalnum() or ch.isspace() or ch == "."
  )
  return _MULTI_SPACE.sub(" ", cleaned).strip()


def proper_trim(value: Any) -> Any:
  """TRIM + remove symbols + PROPER case for text; leave numbers/dates unchanged."""
  if value is None:
    return None
  if isinstance(value, bool):
    return value
  if isinstance(value, (int, float)):
    return value
  if isinstance(value, (datetime, date)):
    return value
  text = remove_symbols(str(value))
  if not text:
    return None
  return text.title()


def apply_text_trim(cell_value: Any, trimmed: Any) -> Any:
  """Use trimmed text only when non-empty; never wipe an existing cell."""
  if cell_value is None or cell_value == "":
    return trimmed
  if trimmed is None:
    return cell_value
  return trimmed


def sanitize_cell_value(value: Any) -> Any:
  """Strip symbols from strings; never blank a cell that had real data."""
  if value is None:
    return None
  if isinstance(value, (bool, int, float, datetime, date)):
    return value
  if isinstance(value, str):
    if not value.strip():
      return None
    cleaned = remove_symbols(value)
    return cleaned if cleaned else value
  return value
