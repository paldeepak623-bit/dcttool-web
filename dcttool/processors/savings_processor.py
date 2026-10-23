"""
Savings processor — DCTtool V2.0
Rules specified by user:
  AdmissionNo          : preserve as-is (no changes)
  ProductDescription   : "Saving Deposits" (always)
  AccountNo            : preserve as-is (no changes)
  DepositDate          : preserve as-is (valid date format preserved)
  IntroducerNo         : Blank (None)
  ChequeOption         : "No" (always)
  OperationTypeDesc    : "Single" (always)
  IsOrganisation       : "No" (always)
  RegisterSlNo         : Blank (None)
  Balance              : preserve as-is (no changes)
  LedgerFolioNo        : Blank (None)
  JointAdmissionNo     : Blank (None)
  PacsIDPkey           : pacs_id (text format @)
"""

from datetime import date
from typing import Any
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from dcttool.date_utils import apply_date_column, excel_date_value, parse_date
from dcttool.processors.cross_sheet import normalize_admission


class SavingsProcessorError(Exception):
    """Raised when processing the Savings sheet fails."""
    pass


HEADER_ROW = 1
PRODUCT = "Saving Deposits"
CHEQUE_OPTION = "No"
OPERATION_TYPE = "Single"
IS_ORGANISATION = "No"


def _header_map(ws: Worksheet) -> dict[str, int]:
    headers: dict[str, int] = {}
    for col in range(1, (ws.max_column or 0) + 1):
        val = ws.cell(HEADER_ROW, col).value
        if val is not None and str(val).strip():
            headers[str(val).strip()] = col
    return headers


def _find_col(headers: dict[str, int], *candidates: str) -> int | None:
    norm_headers = {k.strip().lower().replace("_", "").replace(" ", ""): v for k, v in headers.items()}
    for c in candidates:
        key = c.strip().lower().replace("_", "").replace(" ", "")
        if key in norm_headers:
            return norm_headers[key]
    return None


def _set_cell(
    ws: Worksheet,
    row: int,
    col_idx: int | None,
    value: Any,
    number_format: str | None = None,
) -> None:
    if not col_idx:
        return
    cell = ws.cell(row, col_idx)
    cell.value = value
    if number_format is not None:
        cell.number_format = number_format


def process_savings_sheet(
    ws: Worksheet,
    *,
    pacs_id: str,
    membership_dates: dict[str, date] | None = None,
) -> int:
    """
    Process Savings sheet according to DCTtool V2.0 business rules.
    Syncs DepositDate with Membership AdmissionDate (Membership AdmissionNo = Savings AccountNo).
    Returns processed row count.
    """
    if not pacs_id or not str(pacs_id).strip():
        raise SavingsProcessorError("PacsID is required.")

    headers = _header_map(ws)
    if not headers:
        raise SavingsProcessorError("No headers found on the Savings sheet.")

    pacs_clean = str(pacs_id).strip()

    # Locate column indices
    col_adm = _find_col(headers, "AdmissionNo", "AdmNo")
    col_prod = _find_col(headers, "ProductDescription", "ProductDesc")
    col_acc = _find_col(headers, "AccountNo", "AccNo")
    col_date = _find_col(headers, "DepositDate", "Date")
    col_intro = _find_col(headers, "IntroducerNo", "Introducer")
    col_cheque = _find_col(headers, "ChequeOption", "Cheque")
    col_op_type = _find_col(headers, "OperationTypeDesc", "OperationType")
    col_is_org = _find_col(headers, "IsOrganisation", "IsOrganization")
    col_reg_sl = _find_col(headers, "RegisterSlNo", "RegSlNo")
    col_bal = _find_col(headers, "Balance")
    col_ledger = _find_col(headers, "LedgerFolioNo", "LedgerFolio")
    col_joint = _find_col(headers, "JointAdmissionNo", "JointAdmNo")
    col_pacs = _find_col(headers, "PacsIDPkey", "PacsIDPKey", "PacsIdPkey", "PacsID", "PacsId")

    max_row = ws.max_row or HEADER_ROW
    processed = 0

    for row in range(HEADER_ROW + 1, max_row + 1):
        # Identify valid row: check AdmissionNo or AccountNo
        adm_val = ws.cell(row, col_adm).value if col_adm else None
        acc_val = ws.cell(row, col_acc).value if col_acc else None

        if (adm_val is None or str(adm_val).strip() == "") and (acc_val is None or str(acc_val).strip() == ""):
            continue

        # Rule 1: AdmissionNo — preserve as-is (no changes)

        # Rule 2: ProductDescription — always "Saving Deposits"
        _set_cell(ws, row, col_prod, PRODUCT)

        # Rule 3: AccountNo — preserve as-is (no changes)

        # Rule 4: DepositDate — sync with Membership.AdmissionDate (Membership AdmissionNo = Savings AccountNo)
        # DepositDate cannot be less than AdmissionDate; copy AdmissionDate if member exists in Membership
        if col_date:
            date_raw = ws.cell(row, col_date).value
            acc_key = normalize_admission(acc_val)
            adm_key = normalize_admission(adm_val)

            mem_date: date | None = None
            if membership_dates:
                if acc_key:
                    mem_date = membership_dates.get(acc_key)
                if mem_date is None and adm_key:
                    mem_date = membership_dates.get(adm_key)

            final_dep_date: date | None = None
            if mem_date is not None:
                final_dep_date = mem_date
            else:
                final_dep_date = parse_date(date_raw)

            if final_dep_date is not None:
                if final_dep_date >= date(2025, 3, 30):
                    final_dep_date = date(2025, 3, 30)
                cell = ws.cell(row, col_date)
                cell.value = excel_date_value(final_dep_date)
                cell.number_format = "dd-mm-yyyy"
            else:
                apply_date_column(ws, row, headers, "DepositDate", date_raw)

        # Rule 5: IntroducerNo — Blank (None)
        _set_cell(ws, row, col_intro, None)

        # Rule 6: ChequeOption — always "No"
        _set_cell(ws, row, col_cheque, CHEQUE_OPTION)

        # Rule 7: OperationTypeDesc — always "Single"
        _set_cell(ws, row, col_op_type, OPERATION_TYPE)

        # Rule 8: IsOrganisation — always "No"
        _set_cell(ws, row, col_is_org, IS_ORGANISATION)

        # Rule 9: RegisterSlNo — Blank (None)
        _set_cell(ws, row, col_reg_sl, None)

        # Rule 10: Balance — preserve as-is (no changes)

        # Rule 11: LedgerFolioNo — Blank (None)
        _set_cell(ws, row, col_ledger, None)

        # Rule 12: JointAdmissionNo — Blank (None)
        _set_cell(ws, row, col_joint, None)

        # Rule 13: PacsIDPkey — pacs_id (text format @)
        _set_cell(ws, row, col_pacs, pacs_clean, number_format="@")

        processed += 1

    return processed
