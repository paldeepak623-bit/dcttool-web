"""
LandDetails processor – fills every row of the LandDetails template.

Rules:
  AdmissionNo          : 8-digit zero-padded text (00000012)
  Surveyno             : preserve as-is
  PassBookNo           : preserve as-is
  Titledeedno          : preserve as-is
  Totalextentland      : preserve as-is
  Cultivableland       : preserve as-is
  Landundercultivation : preserve as-is
  LandvaluePerAcre     : preserve as-is
  DeclaredLandinAcres  : preserve as-is
  DeclaredLandValueAmount : Cultivableland * LandvaluePerAcre (calculated)
  FarmcategoryDesc     : "Owner" (always)
  IrrigationtypeDesc   : "Canal" (always)
  LandtypeDesc         : "Wet" (always)
  SoiltypeDesc         : "Black Soil" (always)
  StateName            : "UTTAR PRADESH" (fixed, capitals)
  DistrictDesc         : from tool input (CAPITALS)
  MandalDesc           : randomly assigned from portal scraped mandals
  VillageName          : randomly assigned from the SAME mandal's villages
  Eastboundary         : blank (clear)
  Northboundary        : blank (clear)
  Southboundary        : blank (clear)
  Westboundary         : blank (clear)
  PacsID               : from tool input (Text @ format)
  TotalextentlandOU1   : 0.00
  TotalextentlandOU2   : 0.00
  CultivablelandOU1    : 0.00
  CultivablelandOU2    : 0.00
  LandundercultivationOU1 : 0.00
  LandundercultivationOU2 : 0.00
  DeclaredLandinAcresOU1  : 0.00
  DeclaredLandinAcresOU2  : 0.00
"""

from __future__ import annotations

import random
import shutil
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from dcttool.portals.unite_land_portal import LandPortalData

SHEET_NAME = "LandDetails"
HEADER_ROW = 1
DECIMAL_FMT = "0.00"

# Fixed values (always these values)
STATE_NAME = "UTTAR PRADESH"
FARM_CATEGORY = "Owner"
IRRIGATION_TYPE = "Canal"
LAND_TYPE = "Wet"
SOIL_TYPE = "Black Soil"

# These columns are preserved exactly as source (never modified)
PRESERVE_COLUMNS = {
    "AdmissionNo",
    "Surveyno",
    "Totalextentland",
    "Cultivableland",
    "Landundercultivation",
    "LandvaluePerAcre",
    "DeclaredLandinAcres",
}

# These columns always get 0.00 (PassBookNo and Titledeedno set to 0.00)
ZERO_DECIMAL_COLUMNS = [
    "PassBookNo",
    "Titledeedno",
    "TotalextentlandOU1",
    "TotalextentlandOU2",
    "CultivablelandOU1",
    "CultivablelandOU2",
    "LandundercultivationOU1",
    "LandundercultivationOU2",
    "DeclaredLandinAcresOU1",
    "DeclaredLandinAcresOU2",
]

# These boundary columns are always blank (cleared)
BLANK_COLUMNS = [
    "Eastboundary",
    "Northboundary",
    "Southboundary",
    "Westboundary",
]


class LandProcessorError(Exception):
    pass


def _header_map(ws: Worksheet) -> dict[str, int]:
    headers: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        raw = ws.cell(HEADER_ROW, col).value
        if raw is None:
            continue
        name = str(raw).strip()
        if name:
            headers[name] = col
    return headers


def _set_cell(
    ws: Worksheet,
    row: int,
    headers: dict[str, int],
    name: str,
    value: Any,
    *,
    number_format: str | None = None,
) -> None:
    col = headers.get(name)
    if not col:
        return
    cell = ws.cell(row, col)
    cell.value = value
    if number_format is not None:
        cell.number_format = number_format


def _to_float(value: Any) -> float | None:
    """Convert any value to float safely."""
    if value is None:
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def _format_admission_no(value: Any) -> str:
    """Format AdmissionNo as 8-digit zero-padded string e.g. 00000012."""
    if value is None:
        return ""
    try:
        num = int(float(str(value).strip()))
        return str(num).zfill(8)
    except (ValueError, TypeError):
        return str(value).strip()


def _build_mandal_village_assignments(
    total_rows: int,
    mandal_village_map: dict[str, list[str]],
    target_mandal: str | None = None,
) -> list[tuple[str, str]]:
    """
    Build (mandal, village) pairs for each row.
    Distributes across all available mandals of the district randomly and non-sequentially,
    ensuring that EVERY village strictly belongs to the assigned mandal without any mismatch.
    """
    valid_mandals = [m for m, vils in mandal_village_map.items() if vils]
    if not valid_mandals:
        return []

    # Distribute across all available district mandals (e.g. 4 to 7 mandals)
    m_list = valid_mandals

    assignments: list[tuple[str, str]] = []
    for _ in range(total_rows):
        mandal = random.choice(m_list)
        village = random.choice(mandal_village_map[mandal])
        assignments.append((mandal, village))

    return assignments


def process_land_sheet(
    ws: Worksheet,
    *,
    pacs_id: str,
    district_desc: str,
    portal_data: LandPortalData | None = None,
    target_mandal: str | None = None,
    membership_admissions: set[str] | None = None,
) -> int:
    if not pacs_id or not pacs_id.strip():
        raise LandProcessorError("PacsID is required.")
    if not district_desc or not district_desc.strip():
        raise LandProcessorError("DistrictDesc is required.")

    headers = _header_map(ws)
    if not headers:
        raise LandProcessorError("No headers found on the LandDetails sheet.")

    pacs_value = pacs_id.strip()
    district_value = district_desc.strip().upper()  # CAPITAL LETTERS

    max_row = ws.max_row
    adm_col_idx = headers.get("AdmissionNo", 1)

    # Step 1: Filter rows against valid membership admissions
    # Any row with no admission or an admission NOT in Membership must be deleted
    rows_to_delete: list[int] = []
    valid_rows_original: list[int] = []

    for row in range(HEADER_ROW + 1, max_row + 1):
        adm_raw = ws.cell(row, adm_col_idx).value
        adm_str = "" if adm_raw is None else str(adm_raw).strip()
        adm_norm = ""
        if adm_str and adm_str.upper() not in {"NA", "N/A", "NONE", "NULL"}:
            if adm_str.endswith(".0"):
                adm_str = adm_str[:-2]
            adm_norm = adm_str.lstrip("0") or "0"

        if not adm_norm:
            rows_to_delete.append(row)
        elif membership_admissions is not None and adm_norm not in membership_admissions:
            rows_to_delete.append(row)
        else:
            valid_rows_original.append(row)

    # Delete non-membership / blank rows in reverse contiguous chunks for maximum speed
    if rows_to_delete:
        intervals: list[list[int]] = []
        for r in rows_to_delete:
            if not intervals or r != intervals[-1][1] + 1:
                intervals.append([r, r])
            else:
                intervals[-1][1] = r
        for start_r, end_r in reversed(intervals):
            ws.delete_rows(start_r, end_r - start_r + 1)

    total_valid = len(valid_rows_original)
    if total_valid == 0:
        return 0

    # Build mandal-village assignments if portal data available
    mandal_village_assignments: list[tuple[str, str]] | None = None
    if portal_data and portal_data.mandal_village_map:
        mandal_village_assignments = _build_mandal_village_assignments(
            total_valid, portal_data.mandal_village_map, target_mandal=target_mandal
        )

    processed = 0

    # After contiguous deletion, valid rows are strictly from (HEADER_ROW + 1) to (HEADER_ROW + total_valid)
    for idx in range(total_valid):
        row = HEADER_ROW + 1 + idx

        # Rule 1: AdmissionNo - 8-digit zero-padded text (00000012)
        adm_raw = ws.cell(row, headers.get("AdmissionNo", 1)).value
        adm_formatted = _format_admission_no(adm_raw)
        if "AdmissionNo" in headers:
            cell = ws.cell(row, headers["AdmissionNo"])
            cell.value = adm_formatted
            cell.number_format = "@"  # Text format

        # Rule 2: DeclaredLandValueAmount = Cultivableland * LandvaluePerAcre
        cultivable = _to_float(ws.cell(row, headers["Cultivableland"]).value) if "Cultivableland" in headers else None
        land_val_per_acre = _to_float(ws.cell(row, headers["LandvaluePerAcre"]).value) if "LandvaluePerAcre" in headers else None
        if cultivable is not None and land_val_per_acre is not None:
            declared_land_value = round(cultivable * land_val_per_acre, 2)
            _set_cell(ws, row, headers, "DeclaredLandValueAmount", declared_land_value, number_format=DECIMAL_FMT)

        # Rule 3: Fixed text columns
        _set_cell(ws, row, headers, "FarmcategoryDesc", FARM_CATEGORY)
        _set_cell(ws, row, headers, "IrrigationtypeDesc", IRRIGATION_TYPE)
        _set_cell(ws, row, headers, "LandtypeDesc", LAND_TYPE)
        _set_cell(ws, row, headers, "SoiltypeDesc", SOIL_TYPE)

        # Rule 4: StateName = "UTTAR PRADESH" (fixed, capitals)
        _set_cell(ws, row, headers, "StateName", STATE_NAME)

        # Rule 5: DistrictDesc from tool input (CAPITALS)
        _set_cell(ws, row, headers, "DistrictDesc", district_value)

        # Rule 6 & 7: MandalDesc + VillageName from portal mapping (random)
        if mandal_village_assignments and idx < len(mandal_village_assignments):
            mandal, village = mandal_village_assignments[idx]
            _set_cell(ws, row, headers, "MandalDesc", mandal)
            _set_cell(ws, row, headers, "VillageName", village)
        else:
            # No portal data - leave blank
            _set_cell(ws, row, headers, "MandalDesc", None)
            _set_cell(ws, row, headers, "VillageName", None)

        # Rule 8: Boundary columns = blank (clear whatever is there)
        for col_name in BLANK_COLUMNS:
            _set_cell(ws, row, headers, col_name, None)

        # Rule 9: All OU columns = 0.00
        for col_name in ZERO_DECIMAL_COLUMNS:
            _set_cell(ws, row, headers, col_name, 0.0, number_format=DECIMAL_FMT)

        # Rule 10: PacsID from tool input (Text format @)
        if "PacsID" in headers:
            cell = ws.cell(row, headers["PacsID"])
            cell.value = pacs_value
            cell.number_format = "@"

        processed += 1

    return processed


def process_land_workbook(
    path: Path,
    *,
    sheet_name: str = SHEET_NAME,
    pacs_id: str,
    district_desc: str,
    portal_data: LandPortalData | None = None,
    target_mandal: str | None = None,
    membership_admissions: set[str] | None = None,
    backup: bool = True,
) -> int:
    if not path.exists():
        raise LandProcessorError(f"File not found: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise LandProcessorError(f"Sheet '{sheet_name}' not found in workbook.")
        ws = wb[sheet_name]
        count = process_land_sheet(
            ws,
            pacs_id=pacs_id,
            district_desc=district_desc,
            portal_data=portal_data,
            target_mandal=target_mandal,
            membership_admissions=membership_admissions,
        )
        if backup:
            backup_path = path.with_suffix(path.suffix + ".bak")
            shutil.copy2(path, backup_path)
        wb.save(path)
        return count
    finally:
        wb.close()


def list_land_sheets(path: Path) -> list[str]:
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()
