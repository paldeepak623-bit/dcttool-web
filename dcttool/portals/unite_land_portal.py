"""Login to Unite portal and fetch ALL mandal + villages from Land Details Add form."""

from __future__ import annotations

import random
import sys
import time
from dataclasses import dataclass, field
from typing import Callable
from urllib.parse import urlparse

try:
    from playwright.sync_api import Page, sync_playwright
except ImportError:
    Page = None
    sync_playwright = None

USER_SEL = "#user"
PASS_SEL = "#pwd"
LOGIN_BTN_SEL = "#btnvalidatelogin"
STATE_SEL = "#StateCode"
DISTRICT_SEL = "#DistrictCode"
MANDAL_SEL = "#MandalCode"
VILLAGE_SEL = "#VillageCode"
LAND_ADD_PATH = (
  "/Membership/CaptureCustomerLandDetails/CaptureCustomerLandDetails"
  "?formid=1004&moduleid=1"
)
LOGGED_IN_MARKER = "menuscreen"
STATE_LABEL = "UTTAR PRADESH"
SKIP_OPTIONS = {"", "--select--", "-- select --", "others", "other", "select"}


class UniteLandPortalError(Exception):
  pass


@dataclass(frozen=True)
class LandPortalData:
  """Holds the full mandal -> village mapping scraped from portal."""
  state_name: str
  district_desc: str
  mandal_village_map: dict  # {mandal_name: [village1, village2, ...]}


def normalize_url(url: str) -> str:
  raw = url.strip()
  if not raw:
    raise UniteLandPortalError("URL is required.")
  if not raw.startswith(("http://", "https://")):
    raw = f"https://{raw}"
  parsed = urlparse(raw)
  if not parsed.netloc:
    raise UniteLandPortalError(f"Invalid URL: {url}")
  return f"{parsed.scheme}://{parsed.netloc}"


def _is_logged_in(page: Page) -> bool:
  return LOGGED_IN_MARKER in page.url.lower()


def _login(page: Page, base_url: str, login_id: str, password: str) -> None:
  page.goto(base_url, wait_until="networkidle", timeout=90000)
  page.wait_for_selector(USER_SEL, state="visible", timeout=30000)
  page.fill(USER_SEL, login_id.strip())
  page.fill(PASS_SEL, password)
  try:
    with page.expect_navigation(timeout=90000, wait_until="domcontentloaded"):
      page.click(LOGIN_BTN_SEL)
  except Exception:
    pass
  time.sleep(2)
  if not _is_logged_in(page):
    raise UniteLandPortalError("Login failed. Check URL, Login ID, or Password.")


def _open_land_add(page: Page, base_url: str) -> None:
  add_url = f"{base_url.rstrip('/')}{LAND_ADD_PATH}"
  page.goto(add_url, wait_until="networkidle", timeout=90000)
  page.wait_for_selector(STATE_SEL, state="visible", timeout=45000)


def _read_select_options(page: Page, selector: str) -> list[str]:
  options = page.evaluate(
    """(sel) => {
      const el = document.querySelector(sel);
      if (!el) return [];
      return [...el.options].map(o => (o.text || '').trim());
    }""",
    selector,
  )
  result: list[str] = []
  seen: set[str] = set()
  for name in options:
    key = name.strip().lower()
    if not name or key in SKIP_OPTIONS:
      continue
    if name not in seen:
      seen.add(name)
      result.append(name)
  return result


def _select_option_by_label(page: Page, selector: str, label: str) -> None:
  page.select_option(selector, label=label)


def _find_district_label(districts: list[str], district_desc: str) -> str:
  target = district_desc.strip().lower()
  if not target:
    raise UniteLandPortalError("DistrictDesc is required for portal mapping.")
  for name in districts:
    if name.strip().lower() == target:
      return name

  aliases = {
    "badaun": "budaun",
    "buduan": "budaun",
    "budaun": "budaun",
    "allahabad": "prayagraj",
    "faizabad": "ayodhya",
    "kashi": "varanasi",
    "banaras": "varanasi",
    "bhadohi": "sant ravidas nagar",
    "sant ravidas nagar": "bhadohi",
    "lakhimpur": "kheri",
    "lakhimpur kheri": "kheri",
    "kasganj": "kanshi ram nagar",
    "kanshi ram nagar": "kasganj",
    "hathras": "mahamaya nagar",
    "mahamaya nagar": "hathras",
    "amroha": "jyotiba phule nagar",
    "jyotiba phule nagar": "amroha",
    "sambhal": "bhim nagar",
    "bhim nagar": "sambhal",
    "hapur": "panchsheel nagar",
    "panchsheel nagar": "hapur",
    "shamli": "prabuddh nagar",
    "prabuddh nagar": "shamli",
  }
  alias_target = aliases.get(target)
  if alias_target:
    for name in districts:
      n_low = name.strip().lower()
      if alias_target == n_low or alias_target in n_low:
        return name

  for name in districts:
    if target in name.strip().lower():
      return name

  import difflib
  close = difflib.get_close_matches(target, [d.strip().lower() for d in districts], n=1, cutoff=0.65)
  if close:
    best = close[0]
    for name in districts:
      if name.strip().lower() == best:
        return name

  sample = ", ".join(districts[:8])
  raise UniteLandPortalError(
    f"District '{district_desc}' not found on portal. "
    f"Check spelling (portal sample: {sample}...)."
  )


def _scrape_all_mandal_villages(
  page: Page,
  mandals: list[str],
  *,
  on_status: Callable[[str], None] | None = None,
) -> dict[str, list[str]]:
  """Scrape ALL mandals and their villages. Skip mandals with no villages."""
  if not mandals:
    raise UniteLandPortalError("No mandals found for the selected district.")

  def status(msg: str) -> None:
    if on_status:
      on_status(msg)

  mandal_village_map: dict[str, list[str]] = {}

  for i, mandal_desc in enumerate(mandals):
    status(f"[{i+1}/{len(mandals)}] Scraping mandal: {mandal_desc}...")
    _select_option_by_label(page, MANDAL_SEL, mandal_desc)
    time.sleep(2)
    villages = _read_select_options(page, VILLAGE_SEL)
    if not villages:
      status(f"  Skipped {mandal_desc} - no valid villages.")
      continue
    mandal_village_map[mandal_desc] = villages
    status(f"  {mandal_desc}: {len(villages)} village(s) found.")

  if not mandal_village_map:
    raise UniteLandPortalError(
      "No mandal found with valid villages. "
      "Every mandal had only --Select-- or empty Village list."
    )

  return mandal_village_map


def fetch_land_portal_data(
  url: str,
  login_id: str,
  password: str,
  district_desc: str,
  *,
  headless: bool = True,
  on_status: Callable[[str], None] | None = None,
) -> LandPortalData:
  """Login, open Land Details Add form, return full mandal->village mapping."""
  if not login_id.strip():
    raise UniteLandPortalError("Login ID is required.")

  base_url = normalize_url(url)

  def status(msg: str) -> None:
    if on_status:
      on_status(msg)

  status("Opening Chrome and logging in...")
  with sync_playwright() as p:
    browser = p.chromium.launch(
      channel="chrome",
      headless=headless,
      args=["--disable-popup-blocking"],
    )
    context = browser.new_context(no_viewport=True)
    page = context.new_page()
    try:
      _login(page, base_url, login_id, password)
      status("Login OK - opening Membership > Land Details > Add...")
      _open_land_add(page, base_url)

      _select_option_by_label(page, STATE_SEL, STATE_LABEL)
      time.sleep(2)

      districts = _read_select_options(page, DISTRICT_SEL)
      district_label = _find_district_label(districts, district_desc)
      status(f"Selecting district: {district_label}")
      _select_option_by_label(page, DISTRICT_SEL, district_label)
      time.sleep(2)

      mandals = _read_select_options(page, MANDAL_SEL)
      status(f"Found {len(mandals)} mandals. Scraping all villages...")

      mandal_village_map = _scrape_all_mandal_villages(
        page, mandals, on_status=on_status
      )

      total_villages = sum(len(v) for v in mandal_village_map.values())
      status(
        f"Portal OK - {len(mandal_village_map)} mandal(s), "
        f"{total_villages} total village(s) scraped."
      )
      return LandPortalData(
        state_name=STATE_LABEL,
        district_desc=district_label,
        mandal_village_map=mandal_village_map,
      )
    finally:
      browser.close()


if __name__ == "__main__":
  url = sys.argv[1] if len(sys.argv) > 1 else "up3.unitetools.in"
  login = sys.argv[2] if len(sys.argv) > 2 else ""
  district = sys.argv[3] if len(sys.argv) > 3 else ""
  pwd = sys.argv[4] if len(sys.argv) > 4 else "Unite@123"
  if not login or not district:
    print("Usage: py -3 unite_land_portal.py <url> <login_id> <district> [password]")
    sys.exit(1)
  data = fetch_land_portal_data(
    url, login, pwd, district, headless=True, on_status=print
  )
  print(f"State: {data.state_name}")
  print(f"District: {data.district_desc}")
  for mandal, villages in data.mandal_village_map.items():
    print(f"  {mandal}: {len(villages)} villages")
