"""Shared date parsing and Excel DD-MM-YYYY formatting for all templates."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

DATE_NUMBER_FORMAT = "dd-mm-yyyy"
EXCEL_DATE_EPOCH = datetime(1899, 12, 30)

_DATE_HEADER_EXACT = frozenset({
  "dob",
  "dateofbirth",
  "birthdate",
  "gpissuedate",
})


def parse_date(value: Any) -> date | None:
  """Parse Excel serial numbers, datetime, or DD-MM-YYYY text."""
  if value is None:
    return None
  if isinstance(value, datetime):
    return value.date()
  if isinstance(value, date):
    return value
  if isinstance(value, (int, float)):
    try:
      serial = float(value)
      if 1 <= serial <= 200000:
        return (EXCEL_DATE_EPOCH + timedelta(days=serial)).date()
    except (ValueError, OverflowError):
      return None
    return None

  text = str(value).strip()
  if not text:
    return None

  numeric = text.replace(",", "")
  try:
    if numeric.replace(".", "", 1).isdigit():
      serial = float(numeric)
      if 1 <= serial <= 200000:
        return (EXCEL_DATE_EPOCH + timedelta(days=serial)).date()
  except ValueError:
    pass

  for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%m-%d-%Y", "%d-%m-%y", "%d.%m.%Y"):
    try:
      return datetime.strptime(text, fmt).date()
    except ValueError:
      continue

  parts = text.replace("/", "-").replace(".", "-").split("-")
  if len(parts) == 3:
    try:
      d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
      if y < 100:
        y += 2000
      return date(y, m, d)
    except ValueError:
      pass
  return None


def is_date_column(header: str) -> bool:
  h = header.strip().lower().replace("_", "").replace(" ", "")
  if h in _DATE_HEADER_EXACT:
    return True
  return "date" in header.strip().lower()


def excel_date_value(parsed: date) -> datetime:
  return datetime(parsed.year, parsed.month, parsed.day)


def set_date_cell(
  ws: Worksheet,
  row: int,
  headers: dict[str, int],
  name: str,
  parsed: date | None,
) -> None:
  """Write a real Excel date with DD-MM-YYYY display."""
  if parsed is None:
    return
  col = headers.get(name)
  if not col:
    return
  cell = ws.cell(row, col)
  cell.value = excel_date_value(parsed)
  cell.number_format = DATE_NUMBER_FORMAT


def apply_date_column(
  ws: Worksheet,
  row: int,
  headers: dict[str, int],
  name: str,
  raw: Any,
) -> None:
  """Format date as DD-MM-YYYY; if parse fails keep original value (never blank)."""
  if raw is None or (isinstance(raw, str) and not raw.strip()):
    return
  col = headers.get(name)
  if not col:
    return
  parsed = parse_date(raw)
  if parsed is not None:
    set_date_cell(ws, row, headers, name, parsed)
    return
  cell = ws.cell(row, col)
  cell.value = raw
  if isinstance(raw, (datetime, date)):
    cell.number_format = DATE_NUMBER_FORMAT


def export_date_value(value: Any) -> tuple[Any, str | None]:
  """Convert any date-like value for export; never drop unparseable values."""
  if value is None or (isinstance(value, str) and not str(value).strip()):
    return value, None
  parsed = parse_date(value)
  if parsed is None:
    return value, None
  return excel_date_value(parsed), DATE_NUMBER_FORMAT
