from __future__ import annotations
from difflib import SequenceMatcher
"""Apply all 28 standard business rules to the Membership Excel sheet."""


import random
import re
import shutil
import zlib
from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from dcttool.date_utils import apply_date_column, parse_date, set_date_cell
from dcttool.processors.cross_sheet import normalize_admission, to_float
from dcttool.text_utils import apply_text_trim, proper_trim

SHEET_NAME = "Membership"
HEADER_ROW = 1

# Approved 8 Castes for Hindu
HINDU_CASTES = [
    "Backward Caste A",
    "Backward Caste B",
    "Backward Caste C",
    "Backward Caste D",
    "Other Backward Caste",
    "General",
    "Schedule Caste",
    "Schedule Tribe",
]
CASTE_OPTIONS = HINDU_CASTES

FARMER_OPTIONS = ["Big", "Medium"]

HINDU_SURNAME_MARKERS = {
    "singh", "sharma", "verma", "gupta", "yadav", "kumar", "ram", "lal",
    "pal", "devi", "rani", "prasad", "pandey", "tiwari", "mishra",
    "chaturvedi", "dwivedi", "patel", "chauhan", "lodhi", "saini",
    "maurya", "kashyap", "rathore", "gangwar", "rawat", "joshi",
    "bhatt", "saxena", "dubey", "shukla", "dixit", "bajpai", "awasthi",
    "tripathi", "upadhyay", "ojha", "chaubey", "pandit", "thakur",
    "tomar", "sisodia", "bhardwaj", "goswami", "giri", "puri", "bharti",
    "das", "nath", "babu", "kushwaha", "bind", "prajapati", "nishad"
}

MUSLIM_NAME_MARKERS_EXACT = {
    "ansari", "pathan", "qureshi", "saifi", "iddrisi", "idrisi",
    "siddiqui", "mansoori", "abbasi", "mirza", "syed", "sayeed",
    "ahmad", "ahmed", "mohammad", "mohammed", "muhammad", "mustafa",
    "hussain", "husain", "hasan", "haider", "begum", "bano",
    "khatoon", "fatima", "fathima", "akhtar", "anwar", "iqbal",
    "aslam", "tariq", "rashid", "saleem", "salim", "nazar", "shafiq"
}

NUMERIC_COLUMNS = {
    "AdmissionNo",
    "ShareBalance",
    "ThriftBalance",
    "DividentBalance",
    "Age",
    "BranchId",
    "GroupMemberCount",
    "YeshashiviNo",
    "GPAReferenceAdmissionNo",
    "ContactNo",
    "AdhaarCardNo",
    "AadhaarCardNo",
    "AadharCardNo",
    "DCCBSBACNO",
}

DATE_COLUMNS = {"DOB", "AdmissionDate", "GPAIssueDate"}


class MembershipProcessorError(Exception):
    pass


def is_na(value: Any) -> bool:
    if value is None or isinstance(value, (int, float, bool)):
        return True
    text = str(value).strip().upper()
    return text in {
        "", "0", "00", "000", "0.0", "NA", "N/A", "N.A.", "N.A",
        "-", "--", "---", ".", "..", "...", "?", "??", "NIL", "NONE",
        "NULL", "UNKNOWN", "NOT AVAILABLE"
    }


def is_valid_name(value: Any) -> bool:
    if is_na(value):
        return False
    text = str(value).strip()
    # Check for any digits
    if any(ch.isdigit() for ch in text):
        return False
    # Check for minimum alphabetic characters (at least 2 letters, rejecting single chars like 'A')
    letters = [ch for ch in text if ch.isalpha()]
    if len(letters) < 2:
        return False
    return True


def clean_person_name(value: Any) -> str | None:
    if not is_valid_name(value):
        return None
    trimmed = proper_trim(value)
    if not trimmed or not is_valid_name(trimmed):
        return None
    cleaned = re.sub(r'[^a-zA-Z\s\.]', ' ', str(trimmed))
    cleaned = " ".join(cleaned.split()).strip()
    if len([ch for ch in cleaned if ch.isalpha()]) < 2:
        return None
    return cleaned.title()


def calc_age_years(dob: date | None, admission: date | None) -> int | None:
    """Calculate age in years from DOB to AdmissionDate."""
    if not dob or not admission:
        return None
    years = admission.year - dob.year
    if (admission.month, admission.day) < (dob.month, dob.day):
        years -= 1
    return years


def parse_age(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return None
    if isinstance(value, (int, float)):
        age = int(value)
        if 0 < age < 150:
            return age
        return None
    text = str(value).strip()
    if not text or text.startswith("="):
        return None
    try:
        age = int(float(text))
        if 0 < age < 150:
            return age
    except ValueError:
        pass
    return None


def resolve_age(
    existing: Any,
    dob: date | None,
    admission: date | None,
) -> int:
    """
    Calculate Age:
    1. If DOB & AdmissionDate present -> calculate exact years difference.
    2. If calculated age <= 18 -> set to 21 (portal compliance).
    3. If dates missing -> parse existing age from sheet.
    4. Fallback -> 25.
    """
    calc = calc_age_years(dob, admission)
    if calc is not None:
        if calc <= 18:
            return 21
        return calc

    parsed = parse_age(existing)
    if parsed is not None:
        if parsed <= 18:
            return 21
        return parsed

    return 25


def resolve_community_and_caste(
    existing_community_raw: Any,
    existing_caste_raw: Any,
    member_name: Any,
    surname: Any,
    middle: Any,
    father_name: Any,
    spouse_name: Any,
    rng: random.Random,
) -> tuple[str, str]:
    """
    Strict Rules:
    1. If Community in sheet is Hindu (or Hindu name/caste):
       -> Community = "Hindu"
       -> Caste must be one of:
          ['Backward Caste A', 'Backward Caste B', 'Backward Caste C', 'Backward Caste D',
           'Other Backward Caste', 'General', 'Schedule Caste', 'Schedule Tribe']
       -> NEVER 'Minorities'
    2. If Community in sheet is Muslim (or genuine Muslim name):
       -> Community = "Muslim"
       -> Caste MUST BE ONLY 'Minorities'
    """
    comm_clean = str(existing_community_raw or "").strip().lower()
    caste_clean = str(existing_caste_raw or "").strip().lower()

    matched_hindu_caste = None
    for hc in HINDU_CASTES:
        if caste_clean == hc.lower():
            matched_hindu_caste = hc
            break

    # Determine Community:
    # 1. Sheet explicitly says Hindu
    if "hindu" in comm_clean:
        community = "Hindu"
    # 2. Sheet explicitly says Muslim
    elif "musl" in comm_clean or "islam" in comm_clean:
        community = "Muslim"
    # 3. Sheet has Hindu caste (e.g. General, Backward Caste A, etc.)
    elif matched_hindu_caste is not None:
        community = "Hindu"
    # 4. Sheet has Caste = Minorities
    elif "minor" in caste_clean:
        community = "Muslim"
    else:
        # If community is completely blank in the sheet, check words
        full_text = " ".join(str(p or "") for p in (member_name, surname, middle, father_name, spouse_name)).lower()
        words = set(re.findall(r"[a-z]+", full_text))

        has_hindu_marker = any(w in HINDU_SURNAME_MARKERS for w in words)
        has_muslim_marker = any(w in MUSLIM_NAME_MARKERS_EXACT for w in words)

        # "ali" or "khan" as standalone words
        if ("ali" in words or "khan" in words) and not has_hindu_marker:
            has_muslim_marker = True

        if has_muslim_marker and not has_hindu_marker:
            community = "Muslim"
        else:
            community = "Hindu"

    # Now assign Caste based on resolved Community:
    if community == "Muslim":
        # Rule: aur agar muslim hai to :- kewal Minorities rahega
        caste = "Minorities"
    else:
        # Rule: Hindu :- Backward Caste A, Backward Caste B, Backward Caste C, Backward Caste D,
        #                Other Backward Caste, General, Schedule Caste, Schedule Tribe
        if matched_hindu_caste:
            caste = matched_hindu_caste
        else:
            caste = rng.choice(HINDU_CASTES)

    return community, caste


MEMBERSHIP_HEADER_ALIASES: dict[str, list[str]] = {
    "MemberTypeDescription": [
        "a type", "b type", "membertype", "membertypedescription", 
        "member type", "membertypedesc", "member type description",
        "member_type", "member_type_description"
    ],
    "CustomerTypeDescription": [
        "customertype", "customertypedescription", "customer type", 
        "customer type description", "customer_type", "member/nominal member"
    ],
    "FarmerTypeDescription": [
        "farmertype", "farmertypedescription", "farmer type", 
        "farmer type description", "farmer_type"
    ],
    "AdmissionNo": ["admissionno", "admission number", "admission no", "adm no", "admission_no"],
    "ShareBalance": ["sharebalance", "share balance", "share_balance", "shares"],
    "ThriftBalance": ["thriftbalance", "thrift balance", "thrift_balance"],
    "DividentBalance": ["dividentbalance", "dividend balance", "dividendbalance", "divident balance"],
    "DOB": ["dob", "date of birth", "birth date", "d.o.b."],
    "Age": ["age"],
    "AdmissionDate": ["admissiondate", "admission date", "date of admission", "adm date"],
    "GenderDescription": ["genderdescription", "gender", "gender description", "sex"],
    "MaritalStatusDesc": ["maritalstatusdesc", "maritalstatus", "marital status", "marital status desc", "marital status description"],
    "CommunityDescription": ["communitydescription", "community", "community description", "religion"],
    "CasteDescription": ["castedesc", "caste", "caste description", "caste category", "castedescription"],
    "VillageDescription": ["villagedescription", "village", "village name", "villagename", "village description"],
    "Address1": ["address1", "address 1", "address"],
    "Address2": ["address2", "address 2"],
    "FatherName": ["fathername", "father name", "father's name", "relative name"],
    "SpouseName": ["spousename", "spouse name", "husband name", "wife name"],
    "BranchId": ["branchid", "branch id", "branch"],
    "PacsIDPKey": ["pacsidpkey", "pacs id", "pacsid", "pacs code"],
    "AdhaarCardNo": ["adhaarcardno", "aadhaarcardno", "aadharcardno", "aadhaar", "aadhar", "adhaar"],
}


def _header_map(ws: Worksheet) -> dict[str, int]:
    headers: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        raw = ws.cell(HEADER_ROW, col).value
        if raw is None:
            continue
        name = str(raw).strip()
        if name:
            headers[name] = col

    # Normalize aliases into headers dict
    for canonical, aliases in MEMBERSHIP_HEADER_ALIASES.items():
        if canonical not in headers:
            for alias in aliases:
                for h_name, col_idx in list(headers.items()):
                    if "".join(c for c in h_name.lower() if c.isalnum()) == "".join(c for c in alias.lower() if c.isalnum()):
                        headers[canonical] = col_idx
                        ws.cell(HEADER_ROW, col_idx, value=canonical)
                        break
                if canonical in headers:
                    break

    # Positional guarantees for standard Col 1 and Col 2
    if "CustomerTypeDescription" not in headers and ws.max_column >= 1:
        headers["CustomerTypeDescription"] = 1
        ws.cell(HEADER_ROW, 1, value="CustomerTypeDescription")
    if "MemberTypeDescription" not in headers and ws.max_column >= 2:
        headers["MemberTypeDescription"] = 2
        ws.cell(HEADER_ROW, 2, value="MemberTypeDescription")

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


def filter_village_options(names: list[str]) -> list[str]:
    skip = {"", "--select--", "-- select --", "others", "other", "select", "null", "none", "na", "n/a", "अन्य", "इतर"}
    out: list[str] = []
    seen: set[str] = set()
    for raw in names:
        name = str(raw).strip()
        low = name.lower()
        if not name or low in skip or "other" in low or "select" in low:
            continue
        if name not in seen:
            seen.add(name)
            out.append(name)
    return out


def _clean_tokens(s: str) -> list[str]:
    # Strip common administrative noise words
    noise = r"\b(gram|gaon|vill|village|mouza|majra|purwa|tola|basti|post|tehsil)\b"
    cleaned = re.sub(noise, " ", str(s).lower(), flags=re.IGNORECASE)
    cleaned = re.sub(r"[^a-zA-Z0-9\s]", " ", cleaned)
    return [w for w in cleaned.split() if w]


def _collapse_village(s: str) -> str:
    """Completely collapse string for space-agnostic matching (e.g. 'Rajapurkala' == 'Rajapur Kala')."""
    noise = r"(gram|gaon|vill|village|mouza|majra|purwa|tola|basti)"
    cleaned = re.sub(noise, "", str(s).lower(), flags=re.IGNORECASE)
    return re.sub(r"[^a-zA-Z0-9]", "", cleaned)


def match_pacs_village(
    old_village: str,
    valid_villages: list[str],
    rng: random.Random,
    pacs_name: str = "",
) -> str:
    """Smart-match a member's village to the closest genuine PACS portal village.
    Strictly assigns ONLY villages that exist in the portal dropdown / registered list.
    Eliminates spelling mismatches, handles compound words (Rajapurkala / Rajapur Kala),
    and strictly avoids portal 'Other' / rejection.
    """
    if not valid_villages:
        return str(old_village or "").strip()

    clean_raw = str(old_village or "").strip()
    if not clean_raw or clean_raw.lower() in {"others", "other", "--select--", "null", "none", "na", "n/a", "select"}:
        return rng.choice(valid_villages)

    # 1. Exact case-insensitive match
    for v in valid_villages:
        if clean_raw.lower() == v.strip().lower():
            return v

    # 2. Collapsed alphanumeric match (space-agnostic)
    collapsed_old = _collapse_village(clean_raw)
    if collapsed_old:
        for v in valid_villages:
            if collapsed_old == _collapse_village(v):
                return v

    # 3. Component match for parenthetical / hyphenated names: e.g. "RAJAPURKALA (PYAGPUR)"
    parts = re.split(r"[\(\)\-\,\/]", clean_raw)
    for p in parts:
        c_part = _collapse_village(p)
        if len(c_part) >= 3:
            for v in valid_villages:
                if c_part == _collapse_village(v):
                    return v

    # 4. Substring containment match with length check
    if len(collapsed_old) >= 4:
        best_sub_v = None
        best_sub_len = 0
        for v in valid_villages:
            c_v = _collapse_village(v)
            if len(c_v) >= 4 and (collapsed_old in c_v or c_v in collapsed_old):
                match_len = min(len(c_v), len(collapsed_old))
                if match_len > best_sub_len:
                    best_sub_len = match_len
                    best_sub_v = v
        if best_sub_v:
            return best_sub_v

    # 5. Fuzzy sequence ratio on tokens and collapsed representations
    toks = _clean_tokens(clean_raw)
    old_clean = " ".join(toks) if toks else clean_raw.lower()
    best_v = None
    best_score = 0.0

    for v in valid_villages:
        v_toks = _clean_tokens(v)
        v_clean = " ".join(v_toks)
        
        # Token overlap ratio
        score1 = SequenceMatcher(None, old_clean, v_clean).ratio()
        
        # Collapsed ratio
        score2 = SequenceMatcher(None, collapsed_old, _collapse_village(v)).ratio()
        
        score = max(score1, score2)
        if toks and v_toks and toks[0] == v_toks[0]:
            score = max(score, 0.85)

        if score > best_score:
            best_score = score
            best_v = v

    if best_score >= 0.60 and best_v is not None:
        return best_v

    # 6. Balanced, deterministic distribution across genuine PACS dropdown villages
    # Uses stable crc32 hash of the village string so members with the same raw village stay
    # together in the same genuine PACS village, while spreading all unmatched villages naturally
    # across the genuine PACS portal dropdown options.
    idx = zlib.crc32(clean_raw.lower().encode("utf-8")) % len(valid_villages)
    return valid_villages[idx]



def apply_thrift_transfers(
    ws: Worksheet,
    thrift_by_admission: dict[str, float],
) -> int:
    """Add thrift amounts onto Membership ThriftBalance by AdmissionNo."""
    if not thrift_by_admission:
        return 0
    headers = _header_map(ws)
    adm_col = headers.get("AdmissionNo")
    thrift_col = headers.get("ThriftBalance")
    if not adm_col or not thrift_col:
        return 0
    touched = 0
    for row in range(HEADER_ROW + 1, ws.max_row + 1):
        key = normalize_admission(ws.cell(row, adm_col).value)
        if not key or key not in thrift_by_admission:
            continue
        existing = to_float(ws.cell(row, thrift_col).value) or 0.0
        add = thrift_by_admission[key]
        cell = ws.cell(row, thrift_col)
        cell.value = round(existing + add, 2)
        cell.number_format = "0.00"
        touched += 1
    return touched



def _group_contiguous_rows(rows: list[int]) -> list[tuple[int, int]]:
    """Group sorted row indices into (start_row, count) tuples for delete_rows."""
    if not rows:
        return []
    sorted_rows = sorted(rows)
    ranges = []
    start = sorted_rows[0]
    count = 1
    for r in sorted_rows[1:]:
        if r == start + count:
            count += 1
        else:
            ranges.append((start, count))
            start = r
            count = 1
    ranges.append((start, count))
    return ranges

def process_membership_sheet(
    ws: Worksheet,
    *,
    pacs_id_pkey: str | None = None,
    pacs_name: str | None = None,
    clear_pacs_id: bool = False,
    seed: int | None = None,
    village_options: list[str] | None = None,
    loan_admissions: set[str] | None = None,
    thrift_by_admission: dict[str, float] | None = None,
    on_status: Any = None,
) -> int:
    """Apply all 28 Membership business rules. Returns number of data rows processed."""
    headers = _header_map(ws)
    if not headers:
        raise MembershipProcessorError("No headers found on the Membership sheet.")

    villages = filter_village_options(village_options or [])
    if not villages:
        sheet_vills = []
        v_col = headers.get("VillageDescription") or headers.get("Village") or headers.get("VillageName")
        if v_col:
            for r in range(HEADER_ROW + 1, ws.max_row + 1):
                val = ws.cell(r, v_col).value
                if val and str(val).strip():
                    sheet_vills.append(str(val).strip())
        villages = filter_village_options(sheet_vills)
        if not villages:
            villages = ["Others", "Main Village"]

    rng = random.Random(seed or 42)
    processed = 0
    last_row = ws.max_row
    loan_set = loan_admissions or set()

    # Determine column letters for Age and Aadhaar
    age_col_idx = headers.get("Age", 13)
    aadhaar_col_indices = [
        col_idx for col_name, col_idx in headers.items()
        if any(k in col_name.lower() for k in ("adhaar", "aadhar", "adhar"))
    ]

    # Deletion Rule:
    # If ShareBalance <= 4.0 (i.e. 0, 0.001, 0.01, 0.1, 1, 2, 3, 4, or below 5)
    # AND AdmissionNo is NOT in loan_set (ShortTermLoans, LoanTransaction, LandDetails, Savings)
    # -> Remove row from Membership
    # BUT if AdmissionNo IS in loan_set -> RETAIN member row unconditionally!
    share_col_idx = headers.get("ShareBalance")
    adm_col_idx = headers.get("AdmissionNo")
    if share_col_idx and adm_col_idx:
        rows_to_delete_set: set[int] = set()
        kept_low_share = 0
        for r in range(HEADER_ROW + 1, last_row + 1):
            adm_raw = ws.cell(r, adm_col_idx).value
            if adm_raw is None or str(adm_raw).strip() == "":
                continue
            adm_key = normalize_admission(adm_raw)

            sb_raw = ws.cell(r, share_col_idx).value
            try:
                sb_val = float(sb_raw) if sb_raw is not None else None
            except (ValueError, TypeError):
                sb_val = None

            is_low_share = (sb_val is not None and sb_val <= 4.0) or (
                isinstance(sb_raw, str) and sb_raw.strip() in ("0", "0.0", "0.00", "0.1", "0.01", "0.001", "1", "2", "3", "4")
            )

            if is_low_share:
                if adm_key in loan_set:
                    kept_low_share += 1
                else:
                    rows_to_delete_set.add(r)

        if rows_to_delete_set:
            max_c = ws.max_column
            old_max_r = last_row
            target_r = HEADER_ROW + 1
            for src_r in range(HEADER_ROW + 1, old_max_r + 1):
                if src_r in rows_to_delete_set:
                    continue
                if src_r != target_r:
                    for c in range(1, max_c + 1):
                        src_cell = ws.cell(src_r, c)
                        tgt_cell = ws.cell(target_r, c)
                        tgt_cell.value = src_cell.value
                        tgt_cell.number_format = src_cell.number_format
                target_r += 1
            deleted_count = old_max_r - (target_r - 1)
            if deleted_count > 0:
                ws.delete_rows(target_r, deleted_count)
            last_row = ws.max_row
            if on_status and callable(on_status):
                on_status(f"[OK] Low Share Filter: {len(rows_to_delete_set)} inactive member(s) (share <= Rs. 4) removed. {kept_low_share} retained with active loan/land/savings.")

    # 1. Pre-scan for valid father names and member names to build replacement pools
    names_pool: list[str] = []
    member_names_pool: list[str] = []

    father_col_idx = headers.get("FatherName")
    if father_col_idx:
        for r in range(HEADER_ROW + 1, last_row + 1):
            val = clean_person_name(ws.cell(r, father_col_idx).value)
            if val and len(val) >= 2 and val not in names_pool:
                names_pool.append(val)

    mem_name_col_idx = headers.get("MemberName")
    if mem_name_col_idx:
        for r in range(HEADER_ROW + 1, last_row + 1):
            val = clean_person_name(ws.cell(r, mem_name_col_idx).value)
            if val and len(val) >= 2 and val not in member_names_pool:
                member_names_pool.append(val)

    if not names_pool:
        names_pool = ["Ram Kumar", "Shiv Charan", "Om Prakash", "Rajesh Kumar", "Suresh Chandra", "Mahendra Singh", "Dinesh Kumar", "Ramesh Chandra"]
    if not member_names_pool:
        member_names_pool = names_pool

    # Identify regional column names to blank out
    regional_col_names = [name for name in headers if "regional" in name.lower()]

    max_col = ws.max_column
    max_row = last_row

    # Pre-cache column indices to eliminate repeated dictionary lookups
    adm_col = headers.get("AdmissionNo")
    share_col = headers.get("ShareBalance")
    cust_col = headers.get("CustomerTypeDescription")
    mem_type_col = headers.get("MemberTypeDescription")
    dob_col = headers.get("DOB")
    adm_date_col = headers.get("AdmissionDate")
    gpa_col = headers.get("GPAIssueDate")
    age_col = headers.get("Age")
    marital_col = headers.get("MaritalStatusDesc")
    gender_col = headers.get("GenderDescription")
    father_col = headers.get("FatherName")
    spouse_col = headers.get("SpouseName")
    comm_col = headers.get("CommunityDescription")
    caste_col = headers.get("CasteDescription")
    farmer_col = headers.get("FarmerTypeDescription")
    village_col = headers.get("VillageDescription")
    addr1_col = headers.get("Address1")
    addr2_col = headers.get("Address2")
    branch_col = headers.get("BranchId")
    pacs_pkey_col = headers.get("PacsIDPKey")
    pacs_code_col = headers.get("PacsCode")

    # String columns to trim (only under col 32, skipping numeric/dates)
    trim_cols = [
        col for name, col in headers.items()
        if name not in NUMERIC_COLUMNS and name not in DATE_COLUMNS and col < 32
    ]

    village_cache: dict[str, str] = {}

    for row in range(HEADER_ROW + 1, max_row + 1):
        admission_raw = ws.cell(row, adm_col).value if adm_col else None
        if admission_raw is None or str(admission_raw).strip() == "":
            continue

        # Rule 1: Fast text trim on relevant string cells only
        for col in trim_cols:
            cell = ws.cell(row, col)
            val = cell.value
            if val is not None and isinstance(val, str):
                cell.value = proper_trim(val)

        # Rule 2: ShareBalance & CustomerType / MemberType
        # Share >= 100 -> Member / A Type
        # Share < 100 -> If loan: Member / A Type; Else: Nominal Member / B Type
        share_raw = ws.cell(row, headers["ShareBalance"]).value if "ShareBalance" in headers else None
        share_val = to_float(share_raw) or 0.0
        if "ShareBalance" in headers and share_raw is not None and str(share_raw).strip() != "":
            try:
                _set_cell(ws, row, headers, "ShareBalance", round(float(share_raw), 2), number_format="0.00")
            except (ValueError, TypeError):
                pass
        adm_key = normalize_admission(admission_raw)
        has_loan = adm_key in loan_set

        if share_val >= 100 or has_loan:
            _set_cell(ws, row, headers, "CustomerTypeDescription", "Member")
            _set_cell(ws, row, headers, "MemberTypeDescription", "A Type")
        else:
            _set_cell(ws, row, headers, "CustomerTypeDescription", "Nominal Member")
            _set_cell(ws, row, headers, "MemberTypeDescription", "B Type")

        # Rule 3: Dates (DOB, AdmissionDate, GPAIssueDate formatted DD-MM-YYYY)
        dob_raw = _cell_text(ws, row, headers, "DOB")
        adm_date_raw = _cell_text(ws, row, headers, "AdmissionDate")
        apply_date_column(ws, row, headers, "DOB", dob_raw)

        # Rule 3b: AdmissionDate must never exceed 31-03-2025.
        # If AdmissionDate >= 30-03-2025 (or future date), cap it to 30-03-2025.
        adm_dt = parse_date(adm_date_raw)
        if adm_dt is not None:
            if adm_dt >= date(2025, 3, 30):
                adm_dt = date(2025, 3, 30)
            set_date_cell(ws, row, headers, "AdmissionDate", adm_dt)
        else:
            apply_date_column(ws, row, headers, "AdmissionDate", adm_date_raw)

        gpa_raw = _cell_text(ws, row, headers, "GPAIssueDate")
        apply_date_column(ws, row, headers, "GPAIssueDate", gpa_raw)

        # Rule 4: Age calculation (AdmissionDate - DOB, <= 18 -> 21)
        dob_dt = parse_date(dob_raw)
        existing_age = _cell_text(ws, row, headers, "Age")
        age = resolve_age(existing_age, dob_dt, adm_dt)
        _set_cell(ws, row, headers, "Age", int(age), number_format="0")

        # Rule 5: MaritalStatusDesc -> Married
        _set_cell(ws, row, headers, "MaritalStatusDesc", "Married")

        # Rule 6: Gender / Father / Spouse
        # Male: FatherName populated, SpouseName None
        # Female: SpouseName populated (if father was present and spouse empty, shift Father->Spouse), FatherName None
        gender_raw = str(_cell_text(ws, row, headers, "GenderDescription") or "").strip().title()
        father_name = clean_person_name(_cell_text(ws, row, headers, "FatherName"))
        spouse_name = clean_person_name(_cell_text(ws, row, headers, "SpouseName"))

        if gender_raw == "Female":
            if not spouse_name and father_name:
                spouse_name = father_name
            if not spouse_name:
                spouse_name = rng.choice(names_pool)
            father_name = None
        else:
            # Male or unspecified
            gender_raw = "Male"
            _set_cell(ws, row, headers, "GenderDescription", "Male")
            if not father_name:
                father_name = rng.choice(names_pool)
            spouse_name = None

        _set_cell(ws, row, headers, "FatherName", father_name)
        _set_cell(ws, row, headers, "SpouseName", spouse_name)

        # Rule 7: Blank out all Regional columns
        for reg_col in regional_col_names:
            _set_cell(ws, row, headers, reg_col, None)

        # Clean and sanitize MemberName, MemberSurName, MiddleName
        raw_m_name = _cell_text(ws, row, headers, "MemberName")
        member_name = clean_person_name(raw_m_name)
        if not member_name:
            member_name = rng.choice(member_names_pool)
        _set_cell(ws, row, headers, "MemberName", member_name)

        surname = clean_person_name(_cell_text(ws, row, headers, "MemberSurName"))
        middle = clean_person_name(_cell_text(ws, row, headers, "MiddleName"))
        _set_cell(ws, row, headers, "MemberSurName", surname)
        _set_cell(ws, row, headers, "MiddleName", middle)

        # Rule 8: Community and Caste resolution
        # - If Hindu -> Community="Hindu", Caste is one of the 8 approved Hindu castes (never Minorities)
        # - If Muslim -> Community="Muslim", Caste is ONLY "Minorities"
        existing_comm = _cell_text(ws, row, headers, "CommunityDescription")
        existing_caste = _cell_text(ws, row, headers, "CasteDescription")

        community, caste = resolve_community_and_caste(
            existing_comm,
            existing_caste,
            member_name,
            surname,
            middle,
            father_name,
            spouse_name,
            rng,
        )
        _set_cell(ws, row, headers, "CommunityDescription", community)
        _set_cell(ws, row, headers, "CasteDescription", caste)

        # Rule 9: FarmerTypeDescription -> strictly only "Big" or "Medium"
        existing_farmer = proper_trim(_cell_text(ws, row, headers, "FarmerTypeDescription"))
        if existing_farmer and existing_farmer.title() in ("Big", "Medium"):
            _set_cell(ws, row, headers, "FarmerTypeDescription", existing_farmer.title())
        else:
            _set_cell(ws, row, headers, "FarmerTypeDescription", rng.choice(FARMER_OPTIONS))

        # Rule 10: Village & Address
        # 1. Address1 and Address2 - completely cleared / blank as per user instruction
        _set_cell(ws, row, headers, "Address1", None)
        _set_cell(ws, row, headers, "Address2", None)

        # 2. VillageDescription - wipe old template values and strictly assign ONLY genuine portal village
        old_village = proper_trim(_cell_text(ws, row, headers, "VillageDescription")) or ""
        if old_village in village_cache:
            assigned_village = village_cache[old_village]
        else:
            assigned_village = match_pacs_village(old_village, villages, rng, pacs_name=pacs_name or "")
            village_cache[old_village] = assigned_village
        _set_cell(ws, row, headers, "VillageDescription", assigned_village)

        # Rule 11: BranchId = 101
        _set_cell(ws, row, headers, "BranchId", 101, number_format="0")

        # Rule 12: PacsIDPKey = pacs_id (text format)
        if pacs_id_pkey:
            clean_pacs = str(pacs_id_pkey).strip()
            _set_cell(ws, row, headers, "PacsIDPKey", clean_pacs, number_format="@")
            _set_cell(ws, row, headers, "PacsID", clean_pacs, number_format="@")

        # Rule: Aadhaar Card Numbers must be completely cleared/removed (all blank)
        for c_idx in aadhaar_col_indices:
            cell = ws.cell(row, c_idx)
            if cell.value is not None:
                cell.value = None

        # Rule 13: Fast clearing of columns 32 onward
        for col_idx in range(32, max_col + 1):
            cell = ws.cell(row, col_idx)
            if cell.value is not None:
                cell.value = None

        processed += 1

    # Optional thrift transfers
    if thrift_by_admission:
        apply_thrift_transfers(ws, thrift_by_admission)

    return processed
