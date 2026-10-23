# -*- coding: utf-8 -*-
import os
import sys
import shutil
import zipfile
from pathlib import Path
from flask import Flask, request, jsonify, send_from_directory, send_file
import openpyxl

# Ensure current dir is in sys.path
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from dcttool.master_engine import (
    MasterSettings,
    TemplateSelection,
    process_master_workbook,
    save_split_templates,
    MasterEngineError,
)
from dcttool.sheet_populator import extract_pacs_metadata_from_workbook, count_data_rows
from dcttool.processors.cross_sheet import to_float, header_map, HEADER_ROW

app = Flask(__name__, static_folder="web_ui", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 250 * 1024 * 1024  # 250 MB max file upload

UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Active session state
ACTIVE_SESSION = {
    "file_name": None,
    "file_path": None,
    "file_size": None,
    "detected_district": None,
    "pacs_id": None,
    "audit": None,
    "last_zip": None,
}

@app.route("/")
def index():
    return send_from_directory("web_ui", "index.html")

@app.route("/<path:path>")
def static_files(path):
    if (BASE_DIR / "web_ui" / path).is_file():
        return send_from_directory("web_ui", path)
    if (BASE_DIR / "assets" / path).is_file():
        return send_from_directory("assets", path)
    return "Not Found", 404

@app.route("/api/browse_file", methods=["POST"])
def browse_file():
    # In cloud web environment, instruct browser to open native HTML file picker
    return jsonify({"success": False, "use_html": True})

@app.route("/api/get_audit", methods=["GET"])
def get_audit():
    return jsonify({"audit": ACTIVE_SESSION.get("audit")})

@app.route("/api/remove_file", methods=["POST"])
def remove_file():
    ACTIVE_SESSION["file_name"] = None
    ACTIVE_SESSION["file_path"] = None
    ACTIVE_SESSION["audit"] = None
    return jsonify({"success": True})

@app.route("/api/upload_file", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"success": False, "message": "No file uploaded"}), 400
    f = request.files["file"]
    if not f or not f.filename:
        return jsonify({"success": False, "message": "Empty filename"}), 400

    fname = f.filename
    dest_path = UPLOAD_DIR / fname
    f.save(str(dest_path))
    size_bytes = dest_path.stat().st_size
    size_str = f"{size_bytes / (1024 * 1024):.2f} MB" if size_bytes > 1024 * 1024 else f"{size_bytes / 1024:.1f} KB"

    # Analyze workbook
    try:
        wb = openpyxl.load_workbook(dest_path, data_only=True)
    except Exception as e:
        return jsonify({"success": False, "message": f"Invalid Excel file: {e}"}), 400

    clean_pacs, clean_name, clean_dist = extract_pacs_metadata_from_workbook(wb, fallback_stem=dest_path.stem)

    # Perform audit
    audit = {
        "has_membership": False,
        "membership_count": 0,
        "has_stloans": False,
        "stloans_count": 0,
        "stloans_principal": 0.0,
        "stloans_interest": 0.0,
        "has_lt": False,
        "lt_count": 0,
        "lt_principal": 0.0,
        "lt_interest": 0.0,
        "principal_match": False,
        "interest_match": False,
        "has_land": False,
        "land_count": 0,
        "has_savings": False,
        "savings_count": 0,
        "share_members_count": 0,
    }

    # Membership sheet
    if "Membership" in wb.sheetnames:
        audit["has_membership"] = True
        m_ws = wb["Membership"]
        audit["membership_count"] = count_data_rows(m_ws)
        audit["share_members_count"] = audit["membership_count"]

    # ShortTermLoans sheet
    if "ShortTermLoans" in wb.sheetnames:
        st_ws = wb["ShortTermLoans"]
        audit["has_stloans"] = True
        audit["stloans_count"] = count_data_rows(st_ws)
        hdrs = header_map(st_ws)
        p_col = hdrs.get("OutstandingPrincipal")
        i_col = hdrs.get("OutstandingInterest")
        tot_p = 0.0
        tot_i = 0.0
        for r in range(HEADER_ROW + 1, st_ws.max_row + 1):
            if p_col:
                tot_p += abs(to_float(st_ws.cell(r, p_col).value) or 0.0)
            if i_col:
                tot_i += abs(to_float(st_ws.cell(r, i_col).value) or 0.0)
        audit["stloans_principal"] = round(tot_p, 2)
        audit["stloans_interest"] = round(tot_i, 2)

    # LoanTransaction sheet
    lt_sheet_name = None
    for name in ["LoanTransaction", "LoanTransactions", "OtherLoanTransactions", "ShortTermLoans_To_LT"]:
        if name in wb.sheetnames:
            lt_sheet_name = name
            break

    if lt_sheet_name:
        lt_ws = wb[lt_sheet_name]
        audit["has_lt"] = True
        audit["lt_count"] = count_data_rows(lt_ws)
        l_hdrs = header_map(lt_ws)
        lp_col = l_hdrs.get("BalancePrincipal")
        li_col = l_hdrs.get("BalanceInterest")
        lt_tot_p = 0.0
        lt_tot_i = 0.0
        for r in range(HEADER_ROW + 1, lt_ws.max_row + 1):
            if lp_col:
                lt_tot_p += abs(to_float(lt_ws.cell(r, lp_col).value) or 0.0)
            if li_col:
                lt_tot_i += abs(to_float(lt_ws.cell(r, li_col).value) or 0.0)
        audit["lt_principal"] = round(lt_tot_p, 2)
        audit["lt_interest"] = round(lt_tot_i, 2)

    # LandDetails sheet
    for name in ["LandDetails", "Land Details"]:
        if name in wb.sheetnames:
            audit["has_land"] = True
            audit["land_count"] = count_data_rows(wb[name])
            break

    # Savings sheet
    if "Savings" in wb.sheetnames:
        audit["has_savings"] = True
        audit["savings_count"] = count_data_rows(wb["Savings"])

    # Reconciliation
    if audit["has_stloans"] and audit["has_lt"]:
        audit["principal_match"] = (abs(audit["stloans_principal"] - audit["lt_principal"]) < 0.01)
        audit["interest_match"] = (abs(audit["stloans_interest"] - audit["lt_interest"]) < 0.01)
    else:
        audit["principal_match"] = True
        audit["interest_match"] = True

    ACTIVE_SESSION["file_name"] = fname
    ACTIVE_SESSION["file_path"] = str(dest_path)
    ACTIVE_SESSION["file_size"] = size_str
    ACTIVE_SESSION["detected_district"] = clean_dist or ""
    ACTIVE_SESSION["pacs_id"] = clean_pacs or ""
    ACTIVE_SESSION["audit"] = audit

    wb.close()

    return jsonify({
        "success": True,
        "file_name": fname,
        "file_path": str(dest_path),
        "file_size": size_str,
        "detected_district": clean_dist or "",
        "pacs_id": clean_pacs or "",
        "audit": audit,
    })

@app.route("/api/start_process", methods=["POST"])
def start_process():
    data = request.json or {}
    pacs_id = data.get("pacs_id", "").strip()
    district = data.get("district", "").strip()
    portal_url = data.get("portal_url", "").strip()
    login_id = data.get("login_id", "").strip()
    password = data.get("password", "")
    tpl_states = data.get("templates", {})

    seed_path = ACTIVE_SESSION.get("file_path")
    if not seed_path or not Path(seed_path).exists():
        return jsonify({"success": False, "error": "Master sheet not found or expired. Please upload again."}), 400

    logs: list[str] = []
    def on_status(msg: str):
        logs.append(msg)

    selection = TemplateSelection(
        declaration=tpl_states.get("declaration", True),
        membership=tpl_states.get("membership", True),
        stloans=tpl_states.get("stloans", True),
        loan_transaction=tpl_states.get("loan_transaction", True),
        land=tpl_states.get("land", True),
        savings=tpl_states.get("savings", True),
        stloans_to_lt=tpl_states.get("stloans_to_lt", True),
        dropdown_list=tpl_states.get("dropdown_list", False),
    )

    settings = MasterSettings(
        seed_path=Path(seed_path),
        url=portal_url or "https://up3.unitetools.in/",
        login_id=login_id or "",
        password=password or "",
        pacs_id=pacs_id or ACTIVE_SESSION.get("pacs_id") or "",
        district=district or ACTIVE_SESSION.get("detected_district") or "",
        templates=selection,
    )

    try:
        logs.append(f"⚡ Processing master workbook: {ACTIVE_SESSION.get('file_name')}...")
        res = process_master_workbook(settings, on_status=on_status)

        # Export split templates
        temp_export_dir = OUTPUT_DIR / f"export_{pacs_id or 'pacs'}"
        if temp_export_dir.exists():
            shutil.rmtree(temp_export_dir)
        temp_export_dir.mkdir(parents=True, exist_ok=True)

        exported_paths = save_split_templates(
            res.workbook,
            temp_export_dir,
            selection,
            report=res.report,
            original_path=Path(seed_path),
            on_status=on_status,
        )

        # Create zip archive of all exported templates
        zip_filename = f"DCT_Split_Templates_{pacs_id or 'all'}.zip"
        zip_path = OUTPUT_DIR / zip_filename
        if zip_path.exists():
            zip_path.unlink()

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_f:
            for p in exported_paths:
                zip_f.write(p, arcname=p.name)

        ACTIVE_SESSION["last_zip"] = str(zip_path)
        logs.append(f"✔ Archive created successfully: {zip_filename} ({len(exported_paths)} files)")

        return jsonify({
            "success": True,
            "logs": logs,
            "export_dir": zip_filename,
            "download_url": "/api/download_output",
            "audit": ACTIVE_SESSION.get("audit"),
        })

    except Exception as e:
        import traceback
        err_str = f"Error during processing: {str(e)}"
        logs.append(err_str)
        logs.append(traceback.format_exc())
        return jsonify({
            "success": False,
            "error": str(e),
            "logs": logs,
        }), 500

@app.route("/api/download_output", methods=["GET"])
def download_output():
    last_zip = ACTIVE_SESSION.get("last_zip")
    if not last_zip or not Path(last_zip).exists():
        return "No generated templates found to download. Please run the process first.", 404
    return send_file(last_zip, as_attachment=True, download_name=Path(last_zip).name)

@app.route("/api/open_output", methods=["POST"])
def open_output():
    return jsonify({"success": True})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"============================================================")
    print(f"  DCTtool V2.0 Cloud Web Application")
    print(f"  Live Server listening on: http://0.0.0.0:{port}")
    print(f"============================================================")
    app.run(host="0.0.0.0", port=port, debug=False)
