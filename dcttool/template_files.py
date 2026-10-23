"""Bundled portal upload templates — copy/download and path resolution."""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

# Master workbook sheet name -> bundled template filename
UPLOAD_TEMPLATE_FILES: dict[str, str] = {
    "Membership": "MembershipTemplate.xlsx",
    "ShortTermLoans": "STLoansTemplate.xlsx",
    "Shorts terms loans": "STLoansTemplate.xlsx",
    "Short Term Loans": "STLoansTemplate.xlsx",
    "LandDetails": "LandDetailsTemplate.xlsx",
    "Land Details": "LandDetailsTemplate.xlsx",
    "LoanTransaction": "LoanTransactionTemplate.xlsx",
    "Loan Transaction": "LoanTransactionTemplate.xlsx",
    "Savings": "SavingsTemplate.xlsx",
    "ShortTermLoans To LoanTransaction": "LoanTransactionTemplate.xlsx",
}

# When portal template header differs from master sheet header
HEADER_ALIASES: dict[str, list[str]] = {
    "PurposeDescription": ["CropDescription"],
    "MemberTypeDescription": ["A Type", "B Type", "MemberType", "Member Type", "MemberTypeDesc", "Member Type Description", "member_type"],
    "CustomerTypeDescription": ["CustomerType", "Customer Type", "Customer Type Description", "Member/Nominal Member", "customer_type"],
    "FarmerTypeDescription": ["FarmerType", "Farmer Type", "Farmer Type Description", "farmer_type"],
    "CommunityDescription": ["Community", "Community Description", "Religion"],
    "CasteDescription": ["Caste", "Caste Description", "CasteDesc", "Caste Category"],
    "GenderDescription": ["Gender", "Gender Description", "Sex"],
    "MaritalStatusDesc": ["MaritalStatus", "Marital Status", "MaritalStatusDescription", "Marital Status Desc"],
    "VillageDescription": ["Village", "VillageName", "Village Name", "Village Description"],
    "AdhaarCardNo": ["AadhaarCardNo", "AadharCardNo", "Aadhar No", "Aadhaar No", "Adhaar No", "Aadhaar", "Aadhar"],
}


USER_TEMPLATES_DIR = Path.home() / "Downloads" / "Templets"


def _package_root() -> Path:
    return Path(__file__).resolve().parent


def _bundle_root() -> Path | None:
    meipass = getattr(sys, "_MEIPASS", None)
    return Path(meipass) if meipass else None


def template_search_paths(filename: str) -> list[Path]:
    paths: list[Path] = []
    bundle = _bundle_root()
    if bundle is not None:
        paths.append(bundle / "templates" / filename)
    paths.extend([
        _package_root() / "templates" / filename,
        USER_TEMPLATES_DIR / filename,
        Path.home() / "Downloads" / "STLT" / filename,
        Path.home() / "Downloads" / filename,
    ])
    return paths


def resolve_upload_template(sheet_name: str) -> Path | None:
    from dcttool.sheet_populator import normalize_sheet_title
    canon = normalize_sheet_title(sheet_name) or sheet_name
    fname = UPLOAD_TEMPLATE_FILES.get(sheet_name) or UPLOAD_TEMPLATE_FILES.get(canon)
    if not fname:
        return None
    for path in template_search_paths(fname):
        if path.is_file():
            return path
    return None


def list_bundled_templates() -> list[tuple[str, str, Path]]:
  """Return (sheet_name, filename, path) for each template found in the app bundle."""
  found: list[tuple[str, str, Path]] = []
  for sheet_name, fname in UPLOAD_TEMPLATE_FILES.items():
    path = resolve_upload_template(sheet_name)
    if path is not None:
      found.append((sheet_name, fname, path))
  return found


def copy_templates_to_folder(dest_dir: Path) -> list[Path]:
  """Copy all bundled portal templates to dest_dir (for user download)."""
  dest_dir.mkdir(parents=True, exist_ok=True)
  copied: list[Path] = []
  seen: set[str] = set()
  for _sheet, fname, src in list_bundled_templates():
    if fname in seen:
      continue
    seen.add(fname)
    out = dest_dir / fname
    shutil.copy2(src, out)
    copied.append(out)
  if not copied:
    raise FileNotFoundError("No portal templates found inside the application.")
  return copied
