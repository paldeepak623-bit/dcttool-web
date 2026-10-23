"""
LoanTransaction processor — DCTtool V2.0
Rules (as specified by user):
  AdmissionNo          : preserve as-is
  ProductDescription   : "Normal KCC"
  LoanNo               : preserve as-is
  TransactionDate      : DD-MM-YYYY format
  LedgerFolioNo        : preserve as-is
  DisbursementAmount   : preserve as-is
  OtherChargesDebitAmount : 0.00
  CollectedPrincipal   : DisbursementAmount - BalancePrincipal
  CollectedInterest    : 0.00
  CollectedPenalInterest: 0.00
  CollectedIOD         : 0.00
  CollectedOthers      : 0.00
  BalancePrincipal     : preserve as-is
  BalanceInterest      : preserve as-is
  BalancePenalInterest : 0.00
  BalanceIOD           : 0.00
  Charges              : 0.00
  PacsId               : from tool input (Text @ format)
  BranchId             : 101 (fixed)
  CropDescription/PurposeDescription : "Paddy"
"""

from __future__ import annotations

import shutil
from pathlib import Path
from datetime import date
from typing import Any

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter

from dcttool.date_utils import apply_date_column, set_date_cell, parse_date
from dcttool.processors.cross_sheet import (
    normalize_admission,
    to_float,
)

HEADER_ROW = 1
PRODUCT = "Normal KCC"
CROP_DEFAULT = "Paddy"
BRANCH_ID = 101
DECIMAL_FMT = "0.00"

# These columns get 0.00 always
ZERO_DECIMAL_COLUMNS = [
    "OtherChargesDebitAmount",
    "CollectedInterest",
    "CollectedPenalInterest",
    "CollectedIOD",
    "CollectedOthers",
    "BalancePenalInterest",
    "BalanceIOD",
    "Charges",
]

# These columns are preserved as-is from source (never modified)
PRESERVE_COLUMNS = {
    "AdmissionNo",
    "LoanNo",
    "LedgerFolioNo",
    "DisbursementAmount",
    "BalancePrincipal",
    "BalanceInterest",
}


class LoanProcessorError(Exception):
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


def _cell_val(ws: Worksheet, row: int, headers: dict[str, int], name: str) -> Any:
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


def process_loan_sheet(
    ws: Worksheet,
    *,
    pacs_id: str,
    crop_by_key: dict[str, str] | None = None,
    thrift_already: set[str] | None = None,
    zero_fix_by_loan: dict[str, float] | None = None,
    sanctioned_date_by_admission: dict[str, date] | None = None,
) -> tuple[int, dict[str, float]]:
    """
    Process LoanTransaction sheet.
    Returns (processed_count, thrift_by_admission).
    """
    if not pacs_id or not pacs_id.strip():
        raise LoanProcessorError("PacsId is required.")

    headers = _header_map(ws)
    if not headers:
        raise LoanProcessorError("No headers found on the LoanTransaction sheet.")

    pacs_value = pacs_id.strip()
    processed = 0
    thrift_by_admission: dict[str, float] = {}

    max_col = ws.max_column
    max_row = ws.max_row
    adm_col_idx = headers.get("AdmissionNo")

    for row in range(HEADER_ROW + 1, max_row + 1):
        adm_raw = ws.cell(row, adm_col_idx).value if adm_col_idx else None
        if adm_raw is None or str(adm_raw).strip() == "":
            continue

        # Rule 1: ProductDescription = "Normal KCC"
        _set_cell(ws, row, headers, "ProductDescription", PRODUCT)

        # Rule 2: TransactionDate -> sync with STLoans SanctionedDate if available
        _adm_key_txn = normalize_admission(adm_raw)
        if sanctioned_date_by_admission and _adm_key_txn and _adm_key_txn in sanctioned_date_by_admission:
            # Use the corrected/final SanctionedDate from STLoans
            _tx_dt = sanctioned_date_by_admission[_adm_key_txn]
            if _tx_dt is not None and _tx_dt >= date(2025, 3, 30):
                _tx_dt = date(2025, 3, 30)
            set_date_cell(ws, row, headers, "TransactionDate", _tx_dt)
        else:
            txn_raw = _cell_val(ws, row, headers, "TransactionDate")
            _tx_dt = parse_date(txn_raw)
            if _tx_dt is not None:
                if _tx_dt >= date(2025, 3, 30):
                    _tx_dt = date(2025, 3, 30)
                set_date_cell(ws, row, headers, "TransactionDate", _tx_dt)
            else:
                apply_date_column(ws, row, headers, "TransactionDate", txn_raw)

        # Rule 3: CropDescription / PurposeDescription = "Paddy"
        _set_cell(ws, row, headers, "CropDescription", CROP_DEFAULT)
        _set_cell(ws, row, headers, "PurposeDescription", CROP_DEFAULT)

        # Rule 4: All ZERO columns = 0.00
        for col_name in ZERO_DECIMAL_COLUMNS:
            _set_cell(ws, row, headers, col_name, 0.0, number_format=DECIMAL_FMT)

        # Rule: DisbursementAmount & BalancePrincipal (sync 1000 for small amounts)
        adm_key = normalize_admission(adm_raw)
        if zero_fix_by_loan and adm_key in zero_fix_by_loan:
            disb_val = 1000.0
            bal_val = 1000.0
            _set_cell(ws, row, headers, "DisbursementAmount", disb_val, number_format=DECIMAL_FMT)
            _set_cell(ws, row, headers, "BalancePrincipal", bal_val, number_format=DECIMAL_FMT)
        else:
            disb_raw = _cell_val(ws, row, headers, "DisbursementAmount")
            bal_raw = _cell_val(ws, row, headers, "BalancePrincipal")
            disb_val = to_float(disb_raw)
            bal_val = to_float(bal_raw)
            if bal_val is not None and bal_val < 0:
                bal_val = 0.0
            if disb_val is not None and bal_val is not None:
                # Audio Rule: If DisbursementAmount < BalancePrincipal (negative),
                # set DisbursementAmount equal to BalancePrincipal so that CollectedPrincipal is 0.00
                if disb_val < bal_val:
                    disb_val = bal_val

            if disb_val is not None:
                _set_cell(ws, row, headers, "DisbursementAmount", disb_val, number_format=DECIMAL_FMT)
            if bal_val is not None:
                _set_cell(ws, row, headers, "BalancePrincipal", bal_val, number_format=DECIMAL_FMT)

        # Rule 5: CollectedPrincipal = DisbursementAmount - BalancePrincipal (PASTE AS VALUE)
        if disb_val is not None and bal_val is not None:
            collected_val = round(max(0.0, disb_val - bal_val), 2)
        else:
            collected_val = 0.0
        _set_cell(ws, row, headers, "CollectedPrincipal", collected_val, number_format=DECIMAL_FMT)

        # BalanceInterest (ensure positive)
        if "BalanceInterest" in headers:
            b_int_raw = _cell_val(ws, row, headers, "BalanceInterest")
            b_int_val = to_float(b_int_raw)
            if b_int_val is not None:
                if b_int_val < 0:
                    b_int_val = 0.0
                _set_cell(ws, row, headers, "BalanceInterest", b_int_val, number_format=DECIMAL_FMT)

        # Rule 6: BranchId = 101
        _set_cell(ws, row, headers, "BranchId", BRANCH_ID, number_format="0")

        # Rule 7: PacsId = from tool input (text format)
        if "PacsId" in headers:
            cell = ws.cell(row, headers["PacsId"])
            cell.value = pacs_value
            cell.number_format = "@"

        processed += 1

    return processed, thrift_by_admission


def process_loan_workbook(
    path: Path,
    *,
    sheet_name: str = "LoanTransaction",
    pacs_id: str,
    backup: bool = True,
) -> int:
    if not path.exists():
        raise LoanProcessorError(f"File not found: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        target = None
        for cand in (sheet_name, "LoanTransaction", "OtherLoanTransactions"):
            if cand in wb.sheetnames:
                target = cand
                break
        if target is None:
            raise LoanProcessorError("LoanTransaction sheet not found in workbook.")
        ws = wb[target]
        zero_fix: dict[str, float] = {}
        if "ShortTermLoans" in wb.sheetnames:
            st_ws = wb["ShortTermLoans"]
            st_hdrs = _header_map(st_ws)
            sanc_c = st_hdrs.get("SanctionedAmount")
            adm_c = st_hdrs.get("AdmissionNo")
            if sanc_c and adm_c:
                for r in range(HEADER_ROW + 1, st_ws.max_row + 1):
                    s_val = to_float(st_ws.cell(r, sanc_c).value)
                    if s_val is not None and 0.0 <= s_val < 1.0:
                        adm_k = normalize_admission(st_ws.cell(r, adm_c).value)
                        if adm_k:
                            zero_fix[adm_k] = 1000.0
        count, _ = process_loan_sheet(ws, pacs_id=pacs_id, zero_fix_by_loan=zero_fix)
        if backup:
            backup_path = path.with_suffix(path.suffix + ".bak")
            shutil.copy2(path, backup_path)
        wb.save(path)
        return count
    finally:
        wb.close()


LT_HEADERS_OFFICIAL = [
    "AdmissionNo",
    "ProductDescription",
    "LoanNo",
    "TransactionDate",
    "LedgerFolioNo",
    "DisbursementAmount",
    "OtherChargesDebitAmount",
    "CollectedPrincipal",
    "CollectedInterest",
    "CollectedPenalInterest",
    "CollectedIOD",
    "CollectedOthers",
    "BalancePrincipal",
    "BalanceInterest",
    "BalancePenalInterest",
    "BalanceIOD",
    "Charges",
    "PacsId",
    "BranchId",
    "CropDescription",
]


def convert_stloans_to_loan_transaction(
    wb: openpyxl.workbook.workbook.Workbook,
    pacs_id: str,
    *,
    target_sheet_name: str = "ShortTermLoans_To_LT",
    sanctioned_date_by_admission: dict[str, date] | None = None,
) -> int:
    """
    Generate / Populate LoanTransaction from 'ShortTermLoans' sheet.
    Adheres strictly to user requirements & manual demo:
      1. AdmissionNo          <- ShortTermLoans.AdmissionNo
      2. ProductDescription   <- "Normal KCC"
      3. LoanNo               <- ShortTermLoans.LoanNo
      4. TransactionDate      <- ShortTermLoans.SanctionedDate (DD-MM-YYYY)
      5. LedgerFolioNo        <- None (Blank)
      6. DisbursementAmount   <- SanctionedAmount (< 1.0 => 1000.00) (format 0.00)
      7. OtherChargesDebitAmount <- 0.00
      8. CollectedPrincipal   <- round(max(0.0, DisbursementAmount - BalancePrincipal), 2)
      9. CollectedInterest    <- 0.00
      10. CollectedPenalInterest <- 0.00
      11. CollectedIOD         <- 0.00
      12. CollectedOthers      <- 0.00
      13. BalancePrincipal     <- OutstandingPrincipal (if Sanctioned < 1.0 => 1000.00) (format 0.00)
      14. BalanceInterest      <- OutstandingInterest (format 0.00)
      15. BalancePenalInterest <- 0.00
      16. BalanceIOD           <- 0.00
      17. Charges              <- 0.00
      18. PacsId               <- pacs_id or ShortTermLoans.PacsId (Text @ format)
      19. BranchId             <- ShortTermLoans.BranchId or 101
      20. CropDescription      <- ShortTermLoans.CropDescription or "Paddy"
    """
    if "ShortTermLoans" not in wb.sheetnames:
        raise LoanProcessorError("Cannot generate LoanTransaction: 'ShortTermLoans' sheet not found in workbook.")

    st_ws = wb["ShortTermLoans"]
    st_hdrs = _header_map(st_ws)
    if not st_hdrs:
        raise LoanProcessorError("No headers found on ShortTermLoans sheet.")

    if target_sheet_name in wb.sheetnames:
        lt_ws = wb[target_sheet_name]
        if lt_ws.max_row > 1:
            lt_ws.delete_rows(2, lt_ws.max_row)
    else:
        lt_ws = wb.create_sheet(target_sheet_name)


    for c_idx, h_name in enumerate(LT_HEADERS_OFFICIAL, start=1):
        lt_ws.cell(1, c_idx, value=h_name)

    lt_hdrs = {h: i + 1 for i, h in enumerate(LT_HEADERS_OFFICIAL)}

    adm_c = st_hdrs.get("AdmissionNo")
    loan_c = st_hdrs.get("LoanNo")
    sanc_date_c = st_hdrs.get("SanctionedDate")
    sanc_amt_c = st_hdrs.get("SanctionedAmount")
    out_prin_c = st_hdrs.get("OutstandingPrincipal")
    out_int_c = st_hdrs.get("OutstandingInterest")
    pacs_c = st_hdrs.get("PacsId")
    branch_c = st_hdrs.get("BranchId")
    crop_c = st_hdrs.get("CropDescription")

    pacs_val = str(pacs_id).strip() if pacs_id and str(pacs_id).strip() else None

    out_row = 2
    for r in range(2, st_ws.max_row + 1):
        adm_raw = st_ws.cell(r, adm_c).value if adm_c else None
        if adm_raw is None or str(adm_raw).strip() == "":
            continue

        adm_key = normalize_admission(adm_raw)
        loan_no = st_ws.cell(r, loan_c).value if loan_c else None

        if sanctioned_date_by_admission and adm_key and adm_key in sanctioned_date_by_admission:
            tx_date = sanctioned_date_by_admission[adm_key]
        else:
            tx_date = parse_date(st_ws.cell(r, sanc_date_c).value if sanc_date_c else None)

        if tx_date is not None and tx_date >= date(2025, 3, 30):
            tx_date = date(2025, 3, 30)

        raw_sanc = st_ws.cell(r, sanc_amt_c).value if sanc_amt_c else None
        raw_out_prin = st_ws.cell(r, out_prin_c).value if out_prin_c else None
        raw_out_int = st_ws.cell(r, out_int_c).value if out_int_c else None

        out_prin_val = to_float(raw_out_prin) or 0.0
        out_int_val = to_float(raw_out_int) or 0.0

        if out_prin_val < 0:
            out_prin_val = 0.0
        if out_int_val < 0:
            out_int_val = 0.0

        # Audio rule: if raw_sanc is invalid / REF / letters, replace with out_prin_val
        is_invalid_sanc = False
        if raw_sanc is None or str(raw_sanc).strip() == "":
            is_invalid_sanc = True
        else:
            s_str = str(raw_sanc).strip()
            if "ref" in s_str.lower() or "#" in s_str or any(ch.isalpha() for ch in s_str):
                is_invalid_sanc = True

        if is_invalid_sanc:
            sanc_val = out_prin_val if out_prin_val > 0 else 1000.0
        else:
            sanc_val = to_float(raw_sanc) or 0.0

        if sanc_val < 1.0:
            disb_amt = 1000.0
            bal_prin = 1000.0
        else:
            disb_amt = round(sanc_val, 2)
            bal_prin = round(out_prin_val, 2)

        # Audio Rule: If DisbursementAmount < BalancePrincipal (negative),
        # set DisbursementAmount equal to BalancePrincipal so that CollectedPrincipal is 0.00!
        if disb_amt < bal_prin:
            disb_amt = bal_prin
            coll_prin = 0.0
        else:
            coll_prin = round(disb_amt - bal_prin, 2)

        bal_int = round(out_int_val, 2)

        row_pacs = pacs_val or (str(st_ws.cell(r, pacs_c).value).strip() if pacs_c and st_ws.cell(r, pacs_c).value else "")
        row_branch = st_ws.cell(r, branch_c).value if branch_c and st_ws.cell(r, branch_c).value else BRANCH_ID
        row_crop = st_ws.cell(r, crop_c).value if crop_c and st_ws.cell(r, crop_c).value else CROP_DEFAULT

        # Ensure loan_no fallback to adm_raw if blank or non-digits
        eff_loan_no = loan_no
        if eff_loan_no is None or str(eff_loan_no).strip() == "":
            eff_loan_no = adm_raw
        else:
            l_str = str(eff_loan_no).strip()
            if not l_str.isdigit():
                try:
                    f_l = float(l_str)
                    if f_l.is_integer() and f_l >= 0:
                        eff_loan_no = int(f_l)
                    else:
                        eff_loan_no = adm_raw
                except (ValueError, TypeError):
                    eff_loan_no = adm_raw

        lt_ws.cell(out_row, lt_hdrs["AdmissionNo"], value=adm_raw)
        lt_ws.cell(out_row, lt_hdrs["ProductDescription"], value=PRODUCT)
        lt_ws.cell(out_row, lt_hdrs["LoanNo"], value=eff_loan_no)

        if tx_date:
            set_date_cell(lt_ws, out_row, lt_hdrs, "TransactionDate", tx_date)
        else:
            raw_d = st_ws.cell(r, sanc_date_c).value if sanc_date_c else None
            apply_date_column(lt_ws, out_row, lt_hdrs, "TransactionDate", raw_d)

        lt_ws.cell(out_row, lt_hdrs["LedgerFolioNo"], value=None)

        c = lt_ws.cell(out_row, lt_hdrs["DisbursementAmount"], value=disb_amt)
        c.number_format = DECIMAL_FMT

        c = lt_ws.cell(out_row, lt_hdrs["OtherChargesDebitAmount"], value=0.0)
        c.number_format = DECIMAL_FMT

        c = lt_ws.cell(out_row, lt_hdrs["CollectedPrincipal"], value=coll_prin)
        c.number_format = DECIMAL_FMT

        for z_col in ["CollectedInterest", "CollectedPenalInterest", "CollectedIOD", "CollectedOthers"]:
            c = lt_ws.cell(out_row, lt_hdrs[z_col], value=0.0)
            c.number_format = DECIMAL_FMT

        c = lt_ws.cell(out_row, lt_hdrs["BalancePrincipal"], value=bal_prin)
        c.number_format = DECIMAL_FMT

        c = lt_ws.cell(out_row, lt_hdrs["BalanceInterest"], value=bal_int)
        c.number_format = DECIMAL_FMT

        for z_col in ["BalancePenalInterest", "BalanceIOD", "Charges"]:
            c = lt_ws.cell(out_row, lt_hdrs[z_col], value=0.0)
            c.number_format = DECIMAL_FMT

        c = lt_ws.cell(out_row, lt_hdrs["PacsId"], value=str(row_pacs))
        c.number_format = "@"

        c = lt_ws.cell(out_row, lt_hdrs["BranchId"], value=row_branch)
        c.number_format = "0"

        lt_ws.cell(out_row, lt_hdrs["CropDescription"], value=str(row_crop))

        out_row += 1

    return out_row - 2
