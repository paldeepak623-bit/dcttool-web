"""Apply formatting and business rules to the ShortTermLoans Excel sheet."""

from __future__ import annotations

import shutil
from pathlib import Path
from datetime import date, datetime, timedelta
from typing import Any

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from dcttool.date_utils import apply_date_column, parse_date, set_date_cell
from dcttool.processors.cross_sheet import normalize_admission, to_float
from dcttool.text_utils import apply_text_trim, proper_trim

SHEET_NAME = "ShortTermLoans"
HEADER_ROW = 1

PRODUCT = "Normal KCC"
SCHEME = "SAO"
LOAN_PERIOD = 12
BRANCH_ID = 101
ROI_PERCENTAGE = 7
PENAL_ROI_PERCENTAGE = 2
IOD_PERCENTAGE = 10.7
DECIMAL_FMT = "0.00"

NUMERIC_COLUMNS = {
    "SlNo",
    "AdmissionNo",
    "LoanNo",
    "SanctionedAmount",
    "ROIPercentage",
    "PenalROIPercentage",
    "IODPercentage",
    "OutstandingPrincipal",
    "OutstandingInterest",
    "OutstandingPenalInterest",
    "LoanPeriod",
    "BranchId",
}

DATE_COLUMNS = {"SanctionedDate", "DueDate"}


class STLoansProcessorError(Exception):
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


def _cell_text(ws: Worksheet, row: int, headers: dict[str, int], name: str) -> Any:
    col = headers.get(name)
    if not col:
        return None
    return ws.cell(row, col).value


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


def normalize_crop(value: Any) -> str:
    """
    CropDescription rules:
    - If contains 'Animal' (e.g. Animal Husbandry) -> 'Animals'
    - If contains 'Paddy' -> 'Paddy'
    - Any other value or blank -> 'Paddy'
    - Final value can ONLY be 'Paddy' or 'Animals'
    """
    if value is None:
        return "Paddy"
    text = str(value).strip().lower()
    if not text:
        return "Paddy"
    if "animal" in text:
        return "Animals"
    if "paddy" in text:
        return "Paddy"
    return "Paddy"


def calculate_due_date(sanc_date: date | datetime | str) -> date:
    """
    Calculate DueDate from SanctionedDate:
    Rule: SanctionedDate + 1 Year (12 months) - 2 Days.
    e.g. 01-04-2010 -> 30-03-2011
    """
    if isinstance(sanc_date, str):
        parsed = parse_date(sanc_date)
        if parsed is not None:
            sanc_date = parsed
        else:
            raise ValueError(f"Unable to parse date string: {sanc_date}")
    elif isinstance(sanc_date, datetime):
        sanc_date = sanc_date.date()

    try:
        next_year = sanc_date.year + 1
        if sanc_date.month == 2 and sanc_date.day == 29:
            base_date = date(next_year, 2, 28)
        else:
            base_date = date(next_year, sanc_date.month, sanc_date.day)
        return base_date - timedelta(days=2)
    except Exception:
        return sanc_date + timedelta(days=363)


def process_stloans_sheet(
    ws: Worksheet,
    *,
    pacs_id: str | None = None,
    clear_pacs_id: bool = False,
    membership_dates: dict[str, date] | None = None,
) -> tuple[int, dict[str, float], dict[str, float], dict[str, date]]:
    """
    Apply all 17 rules for ShortTermLoans:
    1. SlNo: 1, 2, 3, 4, 5... sequential
    2. ProductDescription: 'Normal KCC'
    3. AdmissionNo: preserved as is
    4. LoanNo: preserved as is
    5. Scheme: 'SAO'
    6. CropDescription: only 'Paddy' or 'Animals'
    7. SanctionedDate: DD-MM-YYYY (with Date Mapping: adjust if <= MembershipDate)
    8. SanctionedAmount: preserved as is
    9. LoanPeriod: 12
    10. DueDate: DD-MM-YYYY
    11. ROIPercentage: 7
    12. PenalROIPercentage: 2
    13. IODPercentage: 10.7
    14. OutstandingPrincipal: preserved as is
    15. OutstandingInterest: preserved as is
    16. PacsId: pacs_id (text format @)
    17. BranchId: 101
    """
    headers = _header_map(ws)
    if not headers:
        raise STLoansProcessorError("No headers found on the ShortTermLoans sheet.")

    thrift_by_admission: dict[str, float] = {}
    zero_fix_by_loan: dict[str, float] = {}
    final_dates_by_admission: dict[str, date] = {}  # final SanctionedDate per admission
    processed = 0

    clean_pacs = str(pacs_id).strip() if pacs_id else None

    max_col = ws.max_column
    max_row = ws.max_row
    adm_col_idx = headers.get("AdmissionNo")
    loan_col_idx = headers.get("LoanNo")

    # Collect existing numeric admission numbers across the sheet to avoid collisions
    existing_admission_set: set[int] = set()
    if adm_col_idx:
        for r in range(HEADER_ROW + 1, max_row + 1):
            adm_val = ws.cell(r, adm_col_idx).value
            if adm_val is not None:
                norm_str = str(adm_val).strip()
                if norm_str.isdigit():
                    existing_admission_set.add(int(norm_str))
                else:
                    try:
                        f_val = float(norm_str)
                        if f_val.is_integer():
                            existing_admission_set.add(int(f_val))
                    except (ValueError, TypeError):
                        pass

    max_admission = max(existing_admission_set) if existing_admission_set else 0
    next_avail_adm = max(max_admission + 1, 1)
    seen_admissions: set[int] = set()

    for row in range(HEADER_ROW + 1, max_row + 1):
        admission_raw = ws.cell(row, adm_col_idx).value if adm_col_idx else None
        loan_raw = ws.cell(row, loan_col_idx).value if loan_col_idx else None
        if admission_raw is None and loan_raw is None:
            continue

        processed += 1

        # 1. SlNo: 1, 2, 3...
        _set_cell(ws, row, headers, "SlNo", processed, number_format="0")

        # 2. ProductDescription: 'Normal KCC'
        _set_cell(ws, row, headers, "ProductDescription", PRODUCT)

        # 3. AdmissionNo: resolve duplicates by assigning next unique unused number
        current_adm_val = admission_raw
        if adm_col_idx and admission_raw is not None:
            adm_num = None
            adm_str = str(admission_raw).strip()
            if adm_str.isdigit():
                adm_num = int(adm_str)
            else:
                try:
                    f_val = float(adm_str)
                    if f_val.is_integer():
                        adm_num = int(f_val)
                except (ValueError, TypeError):
                    adm_num = None

            if adm_num is not None:
                if adm_num in seen_admissions:
                    # Duplicate detected! Assign unique number not existing in sheet
                    while next_avail_adm in existing_admission_set:
                        next_avail_adm += 1
                    new_adm = next_avail_adm
                    existing_admission_set.add(new_adm)
                    seen_admissions.add(new_adm)
                    next_avail_adm += 1
                    _set_cell(ws, row, headers, "AdmissionNo", new_adm, number_format="0")
                    current_adm_val = new_adm
                else:
                    seen_admissions.add(adm_num)
                    _set_cell(ws, row, headers, "AdmissionNo", adm_num, number_format="0")
                    current_adm_val = adm_num
            else:
                _set_cell(ws, row, headers, "AdmissionNo", admission_raw)

        # 4. LoanNo: if blank, or contains alphabet, or not purely digits -> use AdmissionNo
        use_adm_for_loan = False
        if loan_raw is None or str(loan_raw).strip() == "":
            use_adm_for_loan = True
        else:
            loan_str = str(loan_raw).strip()
            if not loan_str.isdigit():
                try:
                    f_loan = float(loan_str)
                    if f_loan.is_integer() and f_loan >= 0:
                        loan_str = str(int(f_loan))
                    else:
                        use_adm_for_loan = True
                except (ValueError, TypeError):
                    use_adm_for_loan = True

        if use_adm_for_loan:
            _set_cell(ws, row, headers, "LoanNo", current_adm_val, number_format="0")
        else:
            try:
                _set_cell(ws, row, headers, "LoanNo", int(loan_str), number_format="0")
            except (ValueError, TypeError):
                _set_cell(ws, row, headers, "LoanNo", loan_raw)

        # 5. Scheme: 'SAO'
        _set_cell(ws, row, headers, "Scheme", SCHEME)

        # 6. CropDescription: 'Paddy' or 'Animals'
        crop_raw = _cell_text(ws, row, headers, "CropDescription")
        _set_cell(ws, row, headers, "CropDescription", normalize_crop(crop_raw))

        # 7. SanctionedDate: DD-MM-YYYY with Date Mapping & Capping (must not exceed 31-03-2025; >= 30-03-2025 caps to 30-03-2025)
        sanctioned_raw = _cell_text(ws, row, headers, "SanctionedDate")
        adm_key_for_date = normalize_admission(current_adm_val)
        _membership_date: date | None = None
        if membership_dates and adm_key_for_date:
            _membership_date = membership_dates.get(adm_key_for_date)

        final_sanc_date: date | None = None
        candidate: date | None = None

        if _membership_date is not None:
            _sanctioned_parsed = parse_date(sanctioned_raw)
            if _sanctioned_parsed is not None and _sanctioned_parsed > _membership_date:
                # Scenario 1: SanctionedDate already after MembershipDate → keep original
                candidate = _sanctioned_parsed
            else:
                # Scenario 2: SanctionedDate <= MembershipDate → set to MembershipDate + 2 days
                candidate = _membership_date + timedelta(days=2)
        else:
            # No membership date available → parse raw date
            candidate = parse_date(sanctioned_raw)

        if candidate is not None:
            # Capping rule: SanctionedDate must not exceed 31-03-2025; any date >= 30-03-2025 caps to 30-03-2025
            if candidate >= date(2025, 3, 30):
                candidate = date(2025, 3, 30)
            set_date_cell(ws, row, headers, "SanctionedDate", candidate)
            final_sanc_date = candidate
            if adm_key_for_date:
                final_dates_by_admission[adm_key_for_date] = candidate
        else:
            apply_date_column(ws, row, headers, "SanctionedDate", sanctioned_raw)
            final_sanc_date = None

        # Extract Principal & Interest
        p_amt = _cell_text(ws, row, headers, "OutstandingPrincipal")
        p_f = to_float(p_amt)
        i_amt = _cell_text(ws, row, headers, "OutstandingInterest")
        i_f = to_float(i_amt)

        adm_key = normalize_admission(current_adm_val)

        # Rule 14: If OutstandingPrincipal < 0 -> add to ThriftBalance, set ShortTermLoans to 0.00
        if p_f is not None and p_f < 0:
            add_p = round(abs(p_f), 2)
            p_f = 0.0
            if adm_key:
                thrift_by_admission[adm_key] = round(thrift_by_admission.get(adm_key, 0.0) + add_p, 2)

        # Rule 15: If OutstandingInterest < 0 -> add to ThriftBalance, set ShortTermLoans to 0.00
        if i_f is not None and i_f < 0:
            add_i = round(abs(i_f), 2)
            i_f = 0.0
            if adm_key:
                thrift_by_admission[adm_key] = round(thrift_by_admission.get(adm_key, 0.0) + add_i, 2)

        # 8. SanctionedAmount:
        # Audio rule: if invalid/REF/contains alphabets -> replace with OutstandingPrincipal
        sanc_amt = _cell_text(ws, row, headers, "SanctionedAmount")

        is_invalid_sanc = False
        if sanc_amt is None or str(sanc_amt).strip() == "":
            is_invalid_sanc = True
        else:
            sanc_str = str(sanc_amt).strip()
            if "ref" in sanc_str.lower() or "#" in sanc_str:
                is_invalid_sanc = True
            elif any(c.isalpha() for c in sanc_str):
                is_invalid_sanc = True
            elif to_float(sanc_amt) is None and not (isinstance(sanc_amt, str) and sanc_amt.startswith("=")):
                is_invalid_sanc = True

        if is_invalid_sanc:
            if p_f is not None and p_f > 0:
                sanc_f = p_f
            else:
                sanc_f = 1000.0
        else:
            sanc_f = to_float(sanc_amt)
            if sanc_f is None and isinstance(sanc_amt, str) and sanc_amt.startswith("="):
                if p_f is not None:
                    sanc_f = p_f * 2.0 if "*2" in sanc_amt else p_f

        if sanc_f is not None:
            if 0.0 <= sanc_f < 1.0:
                sanc_f = 1000.0
                adm_k = normalize_admission(current_adm_val)
                if adm_k:
                    zero_fix_by_loan[adm_k] = 1000.0
            _set_cell(ws, row, headers, "SanctionedAmount", round(sanc_f, 2), number_format=DECIMAL_FMT)

        # 9. LoanPeriod: 12
        _set_cell(ws, row, headers, "LoanPeriod", LOAN_PERIOD, number_format="0")

        # 10. DueDate: DD-MM-YYYY (Calculated from SanctionedDate: 1 Year - 2 Days)
        if final_sanc_date is not None:
            calc_due = calculate_due_date(final_sanc_date)
            set_date_cell(ws, row, headers, "DueDate", calc_due)
        else:
            due_raw = _cell_text(ws, row, headers, "DueDate")
            apply_date_column(ws, row, headers, "DueDate", due_raw)

        # 11. ROIPercentage: 7
        _set_cell(ws, row, headers, "ROIPercentage", ROI_PERCENTAGE, number_format="0")

        # 12. PenalROIPercentage: 2
        _set_cell(ws, row, headers, "PenalROIPercentage", PENAL_ROI_PERCENTAGE, number_format="0")

        # 13. IODPercentage: 10.7
        _set_cell(ws, row, headers, "IODPercentage", IOD_PERCENTAGE, number_format="0.0")

        # 14. OutstandingPrincipal (now positive)
        if p_f is not None:
            _set_cell(ws, row, headers, "OutstandingPrincipal", p_f, number_format=DECIMAL_FMT)

        # 15. OutstandingInterest (now positive)
        if i_f is not None:
            _set_cell(ws, row, headers, "OutstandingInterest", i_f, number_format=DECIMAL_FMT)

        # OutstandingPenalInterest (blank/None if column exists)
        if "OutstandingPenalInterest" in headers:
            _set_cell(ws, row, headers, "OutstandingPenalInterest", None)

        # 16. PacsId: set from pacs_id (text format @)
        if clean_pacs:
            _set_cell(ws, row, headers, "PacsId", clean_pacs, number_format="@")
        elif clear_pacs_id:
            _set_cell(ws, row, headers, "PacsId", None)

        # 17. BranchId: 101
        _set_cell(ws, row, headers, "BranchId", BRANCH_ID, number_format="0")

    return processed, thrift_by_admission, zero_fix_by_loan, final_dates_by_admission


def process_stloans_workbook(
    path: Path,
    *,
    sheet_name: str = SHEET_NAME,
    pacs_id: str | None = None,
    clear_pacs_id: bool = False,
    backup: bool = True,
) -> int:
    if not path.exists():
        raise STLoansProcessorError(f"File not found: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise STLoansProcessorError(f"Sheet '{sheet_name}' not found in workbook.")
        ws = wb[sheet_name]
        count, _, _, _ = process_stloans_sheet(
            ws,
            pacs_id=pacs_id,
            clear_pacs_id=clear_pacs_id,
        )
        if backup:
            backup_path = path.with_suffix(path.suffix + ".bak")
            shutil.copy2(path, backup_path)
        wb.save(path)
        return count
    finally:
        wb.close()
