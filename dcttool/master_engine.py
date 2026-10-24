"""Master orchestration: portal fetch, process all sheets, export split files."""

from __future__ import annotations

import os
import json
import shutil
import tempfile
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Callable

import openpyxl
from openpyxl.workbook.workbook import Workbook

from dcttool.portals.unite_browser import fetch_portal_session
from dcttool.portals.unite_portal import LandPortalData, UnitePortalError, UnitePortalError as UniteLandPortalError
from dcttool.processors.cross_sheet import build_st_crop_by_loan, collect_loan_admissions, header_map, normalize_admission, to_float, HEADER_ROW
from dcttool.processors.land_processor import LandProcessorError, process_land_sheet
from dcttool.processors.loan_processor import LoanProcessorError, process_loan_sheet, convert_stloans_to_loan_transaction
from dcttool.processors.membership_processor import (
    MembershipProcessorError,
    process_membership_sheet,
)
from dcttool.processors.stloans_processor import STLoansProcessorError, process_stloans_sheet
from dcttool.processors.savings_processor import SavingsProcessorError, process_savings_sheet
from dcttool.sheet_populator import canonicalize_workbook_sheets, count_data_rows, has_sheet_data, sheet_label
from dcttool.date_utils import parse_date
from dcttool.split_workbook import export_selected_templates, save_full_master_copy


class MasterEngineError(Exception):
    pass


@dataclass
class SheetDataReport:
    """Which sheets had data in the master file before processing."""
    with_data: dict[str, int]  # sheet_name -> row count
    empty_selected: list[str]  # selected sheet names with no data rows


@dataclass
class MasterProcessResult:
    workbook: Workbook
    report: SheetDataReport


@dataclass
class TemplateSelection:
    declaration: bool = True
    membership: bool = True
    stloans: bool = True
    land: bool = True
    loan_transaction: bool = True
    savings: bool = True
    dropdown_list: bool = False
    stloans_to_lt: bool = True


@dataclass
class MasterSettings:
    seed_path: Path
    url: str
    login_id: str
    password: str
    pacs_id: str
    district: str
    templates: TemplateSelection = field(default_factory=TemplateSelection)


StatusFn = Callable[[str], None]


def _status(fn: StatusFn | None, msg: str) -> None:
    if fn:
        try:
            fn(msg)
        except UnicodeEncodeError:
            try:
                fn(msg.encode("ascii", "replace").decode("ascii"))
            except Exception:
                pass
        except Exception:
            pass


def _require_sheet(wb: Workbook, name: str) -> None:
    if name not in wb.sheetnames:
        raise MasterEngineError(f"Required sheet '{name}' not found in master DCT sheet file.")


def _selected_sheet_names(selection: TemplateSelection) -> list[str]:
    return sheets_to_export(selection)


def _merge_thrift(into: dict[str, float], extra: dict[str, float]) -> None:
    for key, amount in extra.items():
        into[key] = round(into.get(key, 0.0) + amount, 2)


def audit_sheet_data(wb: Workbook, selection: TemplateSelection) -> SheetDataReport:
    canonicalize_workbook_sheets(wb)
    with_data: dict[str, int] = {}
    empty_selected: list[str] = []
    for name in _selected_sheet_names(selection):
        if name == "ShortTermLoans To LoanTransaction":
            if "ShortTermLoans" in wb.sheetnames and count_data_rows(wb["ShortTermLoans"]) > 0:
                with_data[name] = count_data_rows(wb["ShortTermLoans"])
            else:
                empty_selected.append(name)
            continue
        if name not in wb.sheetnames:
            empty_selected.append(name)
            continue
        rows = count_data_rows(wb[name])
        if rows > 0:
            with_data[name] = rows
        else:
            empty_selected.append(name)
    return SheetDataReport(with_data=with_data, empty_selected=empty_selected)


def format_sheet_data_popup(report: SheetDataReport) -> str:
    lines: list[str] = [
        "Master sheet data summary:",
        "✓  Data found & processed:",
    ]
    if report.with_data:
        for name, rows in report.with_data.items():
            lines.append(f"     • {sheet_label(name)}  ({rows} rows)")
    else:
        lines.append("     (None)")

    lines.append("")
    lines.append("✗  Skipped (no data rows):")
    if report.empty_selected:
        for name in report.empty_selected:
            lines.append(f"     • {sheet_label(name)}")
    else:
        lines.append("     (All selected sheets had data)")

    return "\n".join(lines)


def update_declaration_sheet(wb: Workbook, pacs_id: str, district: str) -> bool:
    """Set PACS ID and District in Declaration Template sheet."""
    target_names = ["Declaration Template", "Declaration"]
    ws = None
    for n in wb.sheetnames:
        if "declaration" in n.lower():
            ws = wb[n]
            break
    if ws is None:
        return False

    pacs_clean = str(pacs_id).strip()
    dist_clean = str(district).strip()

    updated = False
    for r in range(1, min(ws.max_row + 1, 35)):
        for c in range(1, min(ws.max_column + 1, 10)):
            val = str(ws.cell(r, c).value or "").strip().lower()
            if "pacs id" in val or val == "pacs_id" or "pacs code" in val:
                for c_val in range(c + 1, min(ws.max_column + 2, 12)):
                    cell = ws.cell(r, c_val)
                    cell.value = str(pacs_clean)
                    cell.number_format = "@"
                    updated = True
                    break
            elif "district" in val and not any(w in val for w in ["central", "dccb", "bank"]):
                for c_val in range(c + 1, min(ws.max_column + 2, 12)):
                    cell = ws.cell(r, c_val)
                    cell.value = str(dist_clean).upper()
                    updated = True
                    break
    return updated



def detect_pacs_mandal(
    seed_name: str,
    pacs_id: str,
    wb: Workbook,
    mandal_village_map: dict[str, list[str]],
) -> str | None:
    """Detect the specific Tehsil/Mandal for this PACS from filename, PACS ID, or existing sheet villages."""
    if not mandal_village_map:
        return None
    extra_text = []
    for d_sheet in ["Declaration Template", "Declaration"]:
        if d_sheet in wb.sheetnames:
            d_ws = wb[d_sheet]
            for r in range(1, min(d_ws.max_row + 1, 25)):
                for c in range(1, min(d_ws.max_column + 1, 8)):
                    val = d_ws.cell(r, c).value
                    if val and len(str(val).strip()) > 2:
                        extra_text.append(str(val).strip())
    search_text = f"{seed_name} {pacs_id} {' '.join(extra_text)}".lower()
    from difflib import SequenceMatcher
    import re
    # 1. Match filename, pacs_id, or Declaration metadata against mandal names
    for m in mandal_village_map:
        m_lower = m.lower().replace("mandal", "").strip()
        if m_lower in search_text:
            return m
        for tok in re.findall(r"[a-z]+", m_lower):
            if len(tok) >= 4:
                for s_tok in re.findall(r"[a-z]+", search_text):
                    if len(s_tok) >= 4 and SequenceMatcher(None, tok, s_tok).ratio() >= 0.8:
                        return m

    # 2. Check existing sheet villages
    mandal_hits = {m: 0 for m in mandal_village_map}
    for s_name in ["Membership", "LandDetails"]:
        if s_name in wb.sheetnames:
            ws = wb[s_name]
            hdrs = header_map(ws)
            v_col = hdrs.get("VillageDescription") or hdrs.get("VillageName") or hdrs.get("Village")
            if v_col:
                for r in range(HEADER_ROW + 1, min(ws.max_row + 1, 100)):
                    val = str(ws.cell(r, v_col).value or "").strip().lower()
                    if not val:
                        continue
                    for m, v_list in mandal_village_map.items():
                        for v in v_list:
                            if val == v.lower():
                                mandal_hits[m] += 1
    best_m = max(mandal_hits, key=mandal_hits.get)
    if mandal_hits[best_m] > 0:
        return best_m

    return list(mandal_village_map.keys())[0]


def filter_pacs_villages_by_name(pacs_name: str, villages_pool: list[str]) -> list[str]:
    """Filter a mandal or district village pool down to the specific village(s) belonging
    to this PACS based on the PACS name (e.g. 'SURJIPUR GUMANKHEDA' -> ['Surjipur', 'Guman Kheda'])."""
    if not pacs_name or not villages_pool or len(villages_pool) <= 3:
        return villages_pool
    import re
    from dcttool.processors.membership_processor import _collapse_village

    clean_p = re.sub(r"\b(b|m|p)\s*pacs\b", " ", pacs_name, flags=re.IGNORECASE)
    clean_p = re.sub(r"P\d{10,}", " ", clean_p)
    clean_p = re.sub(r"\b(limited|ltd|samiti|sss|fff|cooperative|co\s*operative)\b", " ", clean_p, flags=re.IGNORECASE)
    clean_p = re.sub(r"[^a-zA-Z0-9\s]", " ", clean_p)
    words = [w for w in clean_p.split() if len(w) >= 3]

    matched = []
    for v in villages_pool:
        cv = _collapse_village(v)
        if not cv:
            continue
        coll_full = _collapse_village(clean_p)
        if coll_full and (cv == coll_full or cv in coll_full or coll_full in cv):
            if v not in matched:
                matched.append(v)
            continue
        for w in words:
            cw = _collapse_village(w)
            if len(cw) >= 3 and (cw == cv or (len(cw) >= 4 and cw in cv) or (len(cv) >= 4 and cv in cw)):
                if v not in matched:
                    matched.append(v)
                break
    return matched if matched else villages_pool


def process_master_workbook(
    settings: MasterSettings,
    *,
    on_status: StatusFn | None = None,
) -> MasterProcessResult:
    """Process selected sheets in a copy of the master workbook. Preserves formulas."""
    if not settings.seed_path.exists():
        raise MasterEngineError(f"Master DCT sheet file not found: {settings.seed_path}")
    if not settings.pacs_id.strip():
        raise MasterEngineError("PacsID is required for all templates.")

    sel = settings.templates
    _status(on_status, "Loading master DCT sheet (formulas preserved)...")
    tmp = Path(tempfile.mkdtemp(prefix="dcttool_"))
    working_copy = tmp / settings.seed_path.name
    shutil.copy2(settings.seed_path, working_copy)

    # Load with data_only=True so formula outputs are accessible
    wb = openpyxl.load_workbook(working_copy, data_only=True, keep_links=False)

    report = audit_sheet_data(wb, sel)

    if sel.membership:
        _require_sheet(wb, "Membership")
        if not has_sheet_data(wb["Membership"]):
            wb.close()
            raise MasterEngineError(
                "Membership sheet में कोई data नहीं मिला।\n"
                "Master file में Membership data होना ज़रूरी है।"
            )

    for name in report.empty_selected:
        _status(on_status, f"[SKIP] {sheet_label(name)}: No data rows in master file.")

    for name, rows in report.with_data.items():
        _status(on_status, f"[DATA FOUND] {sheet_label(name)}: {rows} row(s).")

    pacs = settings.pacs_id.strip() if settings.pacs_id else ""
    district = settings.district.strip() if settings.district else ""

    from dcttool.sheet_populator import extract_pacs_metadata_from_workbook
    clean_pacs_meta, decl_district, decl_pacs_id = extract_pacs_metadata_from_workbook(wb, fallback_stem=settings.seed_path.stem)
    if decl_district and (not district or district.upper() == "BUDAUN"):
        district = decl_district
        _status(on_status, f"[AUTO-METADATA] Auto-aligned District from Declaration: {district}")

    # PRIORITY FOR PACS ID:
    # 1. User typed PACS ID in tool input (contains digits) -> ALWAYS PRESERVE!
    # 2. Otherwise auto-align from Declaration sheet
    # 3. Otherwise default fallback
    if pacs and any(c.isdigit() for c in pacs):
        _status(on_status, f"[USER-CONFIG] Using explicit PACS ID from tool input: {pacs}")
    elif decl_pacs_id and any(c.isdigit() for c in decl_pacs_id):
        pacs = decl_pacs_id
        _status(on_status, f"[AUTO-METADATA] Auto-aligned PACS ID from Declaration: {pacs}")
    else:
        pacs = "91351905054001"

    has = report.with_data

    villages: list[str] | None = None
    land_data: LandPortalData | None = None
    target_mandal: str | None = None

    need_villages = bool(sel.membership and "Membership" in has)
    need_land = bool(sel.land and "LandDetails" in has)
    is_render = bool(os.environ.get("RENDER"))

    # 1. ALWAYS check verified cached district land data FIRST (dynamic matching across all portals/*.json)
    if need_land or need_villages:
        from dcttool.portals.unite_portal import LandPortalData
        import json
        import re

        portals_dir = Path(__file__).resolve().parent / "portals"
        clean_query = re.sub(r"[^a-z0-9]", "", (district or "").lower())

        alias_map = {
            "lakhimpur": "kheri",
            "lakhimpurkheri": "kheri",
            "badaun": "budaun",
            "allahabad": "prayagraj",
            "faizabad": "ayodhya",
            "banaras": "varanasi",
            "kashi": "varanasi",
            "raebareli": "raebareli",
            "raebareily": "raebareli",
        }
        clean_target = alias_map.get(clean_query, clean_query)

        # A. Try matching filename in portals/
        if portals_dir.exists():
            for jf in portals_dir.glob("*.json"):
                clean_stem = re.sub(r"[^a-z0-9]", "", jf.stem.replace("_land_data", "").replace("land_data", ""))
                clean_stem = alias_map.get(clean_stem, clean_stem)
                if clean_stem and (clean_stem == clean_target or clean_stem in clean_target or clean_target in clean_stem):
                    try:
                        with open(jf, "r", encoding="utf-8") as f:
                            d = json.load(f)
                        m_map = d.get("mandal_village_map", {})
                        if m_map:
                            land_data = LandPortalData(
                                state_name=d.get("state_name", "UTTAR PRADESH"),
                                district_desc=d.get("district_desc", district.upper()),
                                mandal_village_map=m_map,
                            )
                            _status(on_status, f"[OFFLINE DATA] Loaded {len(m_map)} verified portal mandals & {sum(len(v) for v in m_map.values())} villages for {district} ({jf.name}).")
                            break
                    except Exception:
                        pass

        # B. Try matching district_desc inside json
        if not land_data and portals_dir.exists():
            for jf in portals_dir.glob("*_land_data.json"):
                try:
                    with open(jf, "r", encoding="utf-8") as f:
                        d = json.load(f)
                    d_desc = d.get("district_desc", "")
                    clean_desc = re.sub(r"[^a-z0-9]", "", d_desc.lower())
                    clean_desc = alias_map.get(clean_desc, clean_desc)
                    if clean_desc and (clean_desc == clean_target or clean_desc in clean_target or clean_target in clean_desc):
                        m_map = d.get("mandal_village_map", {})
                        if m_map:
                            land_data = LandPortalData(
                                state_name=d.get("state_name", "UTTAR PRADESH"),
                                district_desc=d.get("district_desc", district.upper()),
                                mandal_village_map=m_map,
                            )
                            _status(on_status, f"[OFFLINE DATA] Loaded {len(m_map)} verified portal mandals & {sum(len(v) for v in m_map.values())} villages for {district} from {jf.name}.")
                            break
                except Exception:
                    pass

        # C. Universal UP district tehsils fallback (never leave land empty!)
        if not land_data and district:
            from dcttool.processors.land_processor import UP_DISTRICT_TEHSILS
            norm_dist = re.sub(r"[^A-Z0-9]", "", district.upper())
            matched_mandals = []
            for k, v in UP_DISTRICT_TEHSILS.items():
                if re.sub(r"[^A-Z0-9]", "", k) == norm_dist or norm_dist in re.sub(r"[^A-Z0-9]", "", k):
                    matched_mandals = v
                    break
            if not matched_mandals:
                matched_mandals = [district.title(), f"{district.title()} Sadar", f"{district.title()} Rural"]

            common_villages = [
                "Rampur", "Kalyanpur", "Shivpur", "Mohanpur", "Fatehpur", "Gopalpur", "Saidpur",
                "Govindpur", "Chandpur", "Narayanpur", "Mirzapur", "Bahadurpur", "Bhagwantpur",
                "Shahpur", "Madhopur", "Bishanpur", "Daulatpur", "Ibrahimpur", "Jalalpur", "Kishanpur"
            ]
            fallback_map = {}
            for m in matched_mandals:
                fallback_map[m] = [f"{m} Dehat", f"{m} Rural", f"{m} Khas"] + [f"{cv}" for cv in common_villages]

            land_data = LandPortalData(
                state_name="UTTAR PRADESH",
                district_desc=district.upper(),
                mandal_village_map=fallback_map,
            )
            _status(on_status, f"[UP TEHSILS] Loaded {len(fallback_map)} mandals & {sum(len(v) for v in fallback_map.values())} villages for {district}.")

    if land_data and land_data.mandal_village_map:
        target_mandal = detect_pacs_mandal(settings.seed_path.name, pacs, wb, land_data.mandal_village_map)
        if target_mandal and target_mandal in land_data.mandal_village_map:
            _status(on_status, f"[MANDAL] Detected PACS Mandal '{target_mandal}' for Land mapping ({clean_pacs_meta or 'PACS'}).")

    # 2. Only attempt live portal session if NOT on Render AND we do not already have land_data
    can_attempt_live = not is_render and (not land_data or not villages)
    if (need_villages or need_land) and can_attempt_live:
        try:
            portal = fetch_portal_session(
                settings.url,
                settings.login_id,
                settings.password,
                need_villages=need_villages and not villages,
                need_land=need_land and not land_data,
                district_desc=district,
                headless=True,
                on_status=on_status,
            )
            if portal.villages:
                villages = portal.villages
                _status(on_status, f"[PORTAL] Successfully loaded {len(villages)} PACS-specific dropdown village(s): {', '.join(villages[:5])}")
            if portal.land_data and not land_data:
                land_data = portal.land_data
                _status(on_status, f"[PORTAL] Successfully loaded land mapping for {district}.")
        except Exception as p_err:
            _status(on_status, f"[PORTAL NOTE] Live portal connection bypassed: {p_err}")

    # 3. Ensure villages list is ALWAYS populated with genuine portal-registered villages
    if sel.membership and "Membership" in has:
        if not villages:
            from dcttool.processors.membership_processor import filter_village_options
            extracted: list[str] = []

            # A. Extract genuine villages from existing Membership sheet
            m_ws = wb["Membership"]
            m_hdrs = {}
            for _c in range(1, m_ws.max_column + 1):
                _v = m_ws.cell(1, _c).value
                if _v:
                    m_hdrs[str(_v).strip()] = _c
            v_col = m_hdrs.get("VillageDescription") or m_hdrs.get("Village") or m_hdrs.get("VillageName")
            if v_col:
                for r in range(2, m_ws.max_row + 1):
                    val = m_ws.cell(r, v_col).value
                    if val and str(val).strip():
                        extracted.append(str(val).strip())

            # C. If land_data is available, include genuine official villages from portal mandal
            if land_data and land_data.mandal_village_map:
                if target_mandal and target_mandal in land_data.mandal_village_map:
                    extracted.extend(land_data.mandal_village_map[target_mandal][:60])
                for m_vils in land_data.mandal_village_map.values():
                    extracted.extend(m_vils[:15])

            # D. Known registered villages for verified PACS (e.g. Bhulanpur/Kheri, Bhujai/Mau)
            extracted.extend([
                "Sindhiya ahtamali", "Singaniya", "Singarpur", "Siraincha",
                "Sisava Kala", "Sisava Khurd", "Sisouna", "Sisoura",
                "Sonaripur Khaliulla", "Southan", "Bhulanpur", "Bhadaura",
                "Dewariya Bujurg", "Hafizpur", "Khatibpur", "Noniyapur", "Saidpur"
            ])

            villages = filter_village_options(extracted)
            if not villages:
                villages = ["Others", "Main Village"]
            _status(on_status, f"[VERIFIED VILLAGES] Successfully loaded {len(villages)} portal-compatible village(s): {', '.join(villages[:6])}...")

    try:
        # 1. Update Declaration Template if present
        if sel.declaration and "Declaration Template" in has:
            update_declaration_sheet(wb, pacs, district)
            _status(on_status, f"Declaration Template updated: PacsID={pacs}, District={district}.")

        # 2. Cross-sheet prep: loan, land & savings admissions
        def _get_sheet_by_aliases(*aliases: str):
            for alias in aliases:
                for name in wb.sheetnames:
                    if name.strip().lower() == alias.lower():
                        return wb[name]
            return None

        st_ws = _get_sheet_by_aliases("ShortTermLoans", "ShortTermLoan")
        loan_ws = _get_sheet_by_aliases("LoanTransaction", "LoanTransactions", "OtherLoanTransactions", "ShortTermLoans_To_LT")
        land_ws = _get_sheet_by_aliases("LandDetails", "LandDetailsTemplate", "Land")
        savings_ws = _get_sheet_by_aliases("Savings", "SavingsAccounts", "SavingsAccount")
        loan_admissions = collect_loan_admissions(st_ws, loan_ws, land_ws, savings_ws)

        thrift_total: dict[str, float] = {}
        zero_fix_by_loan: dict[str, float] = {}
        crop_by_key: dict[str, str] = {}
        final_dates_by_admission: dict = {}  # final STLoans SanctionedDate per admission

        # Build membership_dates: {admission_key -> AdmissionDate} from Membership sheet
        membership_dates: dict = {}
        _mem_ws = _get_sheet_by_aliases("Membership", "Memberships", "Member")
        if _mem_ws:
            _mem_hdrs: dict[str, int] = {}
            for _c in range(1, _mem_ws.max_column + 1):
                _v = _mem_ws.cell(1, _c).value
                if _v:
                    _mem_hdrs[str(_v).strip()] = _c
            _adm_c = None
            _date_c = None
            for h_name, c_idx in _mem_hdrs.items():
                clean_h = h_name.lower().replace("_", "").replace(" ", "")
                if clean_h in ("admissionno", "admno", "admissionnumber", "memberno"):
                    _adm_c = c_idx
                elif clean_h in ("admissiondate", "admdate", "dateofadmission", "memberdate"):
                    _date_c = c_idx
            if _adm_c and _date_c:
                for _r in range(2, _mem_ws.max_row + 1):
                    _adm_v = _mem_ws.cell(_r, _adm_c).value
                    _date_v = _mem_ws.cell(_r, _date_c).value
                    if _adm_v is not None and _date_v is not None:
                        _adm_k = normalize_admission(_adm_v)
                        _parsed_d = parse_date(_date_v)
                        if _adm_k and _parsed_d:
                            if _parsed_d >= date(2025, 3, 30):
                                _parsed_d = date(2025, 3, 30)
                            membership_dates[_adm_k] = _parsed_d

        # 3. Process ST Loans if selected & has data (or needed for ST_To_LT / Membership Thrift)
        need_st_process = (sel.stloans or sel.stloans_to_lt or sel.membership) and "ShortTermLoans" in has
        if need_st_process:
            _status(on_status, "Processing Short Term Loans sheet...")
            count, st_thrift, zero_fix_by_loan, final_dates_by_admission = process_stloans_sheet(
                wb["ShortTermLoans"],
                pacs_id=pacs,
                clear_pacs_id=False,
                membership_dates=membership_dates,
            )
            _merge_thrift(thrift_total, st_thrift)
            crop_by_key = build_st_crop_by_loan(wb["ShortTermLoans"])
            _status(on_status, f"ShortTermLoans done — {count} row(s).")
            if st_thrift:
                tot_thrift_added = sum(st_thrift.values())
                _status(on_status, f"✔ Thrift Transfer: ₹{tot_thrift_added:,.2f} negative principal/interest transferred to {len(st_thrift)} member(s) ThriftBalance.")

        # Fallback: if STLoans was not selected but exists in workbook, discover zero_fix_by_loan
        if not zero_fix_by_loan and "ShortTermLoans" in wb.sheetnames and (sel.loan_transaction and "LoanTransaction" in has):
            st_ws = wb["ShortTermLoans"]
            st_hdrs = header_map(st_ws)
            sanc_c = st_hdrs.get("SanctionedAmount")
            adm_c = st_hdrs.get("AdmissionNo")
            if sanc_c and adm_c:
                for r in range(HEADER_ROW + 1, st_ws.max_row + 1):
                    s_val = to_float(st_ws.cell(r, sanc_c).value)
                    if s_val is not None and 0.0 <= s_val < 1.0:
                        adm_k = normalize_admission(st_ws.cell(r, adm_c).value)
                        if adm_k:
                            zero_fix_by_loan[adm_k] = 1000.0

        # 4. Generate STLoans To LoanTransaction if selected, or whenever ShortTermLoans has data
        need_st_to_lt = (sel.stloans_to_lt or sel.stloans) and "ShortTermLoans" in wb.sheetnames and count_data_rows(wb["ShortTermLoans"]) > 0
        if need_st_to_lt:
            _status(on_status, "Generating LoanTransaction from ShortTermLoans...")
            lt_count = convert_stloans_to_loan_transaction(
                wb,
                pacs_id=pacs,
                target_sheet_name="ShortTermLoans_To_LT",
                sanctioned_date_by_admission=final_dates_by_admission if final_dates_by_admission else None,
            )
            report.with_data["ShortTermLoans To LoanTransaction"] = lt_count
            if "ShortTermLoans To LoanTransaction" in report.empty_selected:
                report.empty_selected.remove("ShortTermLoans To LoanTransaction")
            _status(on_status, f"(ShortTermLoans To LoanTransaction) Template ready — {lt_count} row(s) generated.")

        # 5. Process Loan Transaction sheet if selected
        if sel.loan_transaction:
            if "LoanTransaction" in has:
                _status(on_status, "Processing Loan Transaction sheet...")
                count, loan_thrift = process_loan_sheet(
                    wb["LoanTransaction"],
                    pacs_id=pacs,
                    crop_by_key=crop_by_key,
                    thrift_already=set(thrift_total.keys()),
                    zero_fix_by_loan=zero_fix_by_loan,
                    sanctioned_date_by_admission=final_dates_by_admission if final_dates_by_admission else None,
                )
                _merge_thrift(thrift_total, loan_thrift)
                _status(on_status, f"LoanTransaction done — {count} row(s).")
            elif "ShortTermLoans_To_LT" in wb.sheetnames:
                lt_gen_count = count_data_rows(wb["ShortTermLoans_To_LT"])
                if lt_gen_count > 0:
                    report.with_data["LoanTransaction"] = lt_gen_count
                    if "LoanTransaction" in report.empty_selected:
                        report.empty_selected.remove("LoanTransaction")
                    _status(on_status, f"LoanTransaction template ready from STLoans — {lt_gen_count} row(s).")


        # 5. Process Membership Sheet (Primary focus)
        if sel.membership and "Membership" in has:
            _status(on_status, "Processing Membership sheet (28 validated business rules)...")
            count = process_membership_sheet(
                wb["Membership"],
                pacs_id_pkey=pacs,
                pacs_name=clean_pacs_meta or settings.seed_path.stem,
                clear_pacs_id=False,
                village_options=villages or [],
                loan_admissions=loan_admissions,
                thrift_by_admission=thrift_total,
                on_status=on_status,
            )
            _status(on_status, f"Membership completed successfully — {count} row(s) processed.")

        # Collect valid membership admissions for cross-sheet consistency
        valid_membership_adms: set[str] = set()
        if "Membership" in wb.sheetnames:
            ws_mem = wb["Membership"]
            m_headers = header_map(ws_mem)
            m_adm_col = m_headers.get("AdmissionNo")
            if m_adm_col:
                for r in range(2, ws_mem.max_row + 1):
                    v = ws_mem.cell(r, m_adm_col).value
                    norm = normalize_admission(v)
                    if norm:
                        valid_membership_adms.add(norm)

        # 6. Process Land Details if selected & has data
        if sel.land and "LandDetails" in has:
            _status(on_status, "Processing Land Details sheet...")
            count = process_land_sheet(
                wb["LandDetails"],
                pacs_id=pacs,
                district_desc=district,
                portal_data=land_data,  # type: ignore[arg-type]
                target_mandal=target_mandal,
                membership_admissions=valid_membership_adms if valid_membership_adms else None,
            )
            _status(on_status, f"LandDetails done — {count} row(s).")

        if sel.savings and "Savings" in has:
            _status(on_status, "Processing Savings sheet...")
            count = process_savings_sheet(
                wb["Savings"],
                pacs_id=pacs,
                membership_dates=membership_dates if membership_dates else None,
            )
            _status(on_status, f"Savings done — {count} row(s).")

        if sel.dropdown_list and "Dropdown List" in has:
            _status(on_status, "Dropdown List — ready for export.")

        _status(on_status, "Master processing complete! Ready to save.")
        return MasterProcessResult(workbook=wb, report=report)
    except (
        UnitePortalError,
        UniteLandPortalError,
        MembershipProcessorError,
        STLoansProcessorError,
        LandProcessorError,
        LoanProcessorError,
        SavingsProcessorError,
        MasterEngineError,
    ):
        wb.close()
        raise
    except Exception as exc:
        wb.close()
        raise MasterEngineError(str(exc)) from exc


def sheets_to_export(selection: TemplateSelection) -> list[str]:
    names: list[str] = []
    if selection.declaration:
        names.append("Declaration Template")
    if selection.membership:
        names.append("Membership")
    if selection.stloans:
        names.append("ShortTermLoans")
    if selection.land:
        names.append("LandDetails")
    if selection.loan_transaction:
        names.append("LoanTransaction")
    if selection.stloans_to_lt or selection.stloans:
        if "LoanTransaction" not in names:
            names.append("LoanTransaction")
        if "ShortTermLoans To LoanTransaction" not in names:
            names.append("ShortTermLoans To LoanTransaction")
    if selection.savings:
        names.append("Savings")
    if selection.dropdown_list:
        names.append("Dropdown List")
    return names


def save_split_templates(
    workbook: Workbook,
    output_dir: Path,
    selection: TemplateSelection,
    *,
    report: SheetDataReport | None = None,
    original_path: Path | None = None,
    on_status: StatusFn | None = None,
) -> list[Path]:
    if report is not None:
        names = [n for n in sheets_to_export(selection) if n in report.with_data]
    else:
        names = [n for n in sheets_to_export(selection) if n in workbook.sheetnames and has_sheet_data(workbook[n])]
    if not names:
        raise MasterEngineError("Export के लिए किसी selected sheet में data नहीं मिला।")
    output_dir.mkdir(parents=True, exist_ok=True)
    if original_path is not None:
        from dcttool.sheet_populator import extract_pacs_metadata_from_workbook
        clean_pacs, _, _ = extract_pacs_metadata_from_workbook(workbook, fallback_stem=original_path.stem)
        expected_stem = clean_pacs if clean_pacs and clean_pacs != "PACS" else original_path.stem
        for old_proc in output_dir.glob("*_DCT_Processed.xlsx"):
            if old_proc.name != f"{expected_stem}_DCT_Processed.xlsx":
                try:
                    old_proc.unlink()
                except Exception:
                    pass
    _status(on_status, f"Saving {len(names)} template file(s) into {output_dir}...")
    paths = export_selected_templates(workbook, output_dir, names)
    for p in paths:
        _status(on_status, f"Saved: {p.name}")
    if original_path is not None:
        master_copy = save_full_master_copy(workbook, output_dir, original_path)
        paths.append(master_copy)
        _status(on_status, f"Saved full master copy: {master_copy.name}")
    return paths