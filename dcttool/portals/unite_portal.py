"""Direct HTTP Portal Client for Unite Tools (up3.unitetools.in / up4.unitetools.in).

Provides fast, browser-less login and automatic fetching of Districts, Mandals, and Villages.
"""

from __future__ import annotations

import json
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from typing import Callable
from urllib.parse import urlparse

import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
logger = logging.getLogger(__name__)

SKIP_VILLAGE_NAMES = {"", "--select--", "-- select --", "others", "other", "select", "null", "none", "na", "n/a", "अन्य", "इतर"}


def is_invalid_village(name: str) -> bool:
    s = (name or "").strip().lower()
    if not s or s in SKIP_VILLAGE_NAMES:
        return True
    if "other" in s or "select" in s or "अन्य" in s or s in {"na", "n/a", "none", "null"}:
        return True
    return False
STATE_UP_ID = 33  # Uttar Pradesh


class UnitePortalError(Exception):
    """Raised when portal operations fail."""
    pass


@dataclass(frozen=True)
class LandPortalData:
    """Holds the full mandal -> village mapping scraped from portal."""
    state_name: str
    district_desc: str
    mandal_village_map: dict  # {mandal_name: [village1, village2, ...]}


def normalize_url(url: str) -> str:
    raw = (url or "").strip()
    if not raw:
        raise UnitePortalError("URL is required.")
    if not raw.startswith(("http://", "https://")):
        raw = f"https://{raw}"
    parsed = urlparse(raw)
    if not parsed.netloc:
        raise UnitePortalError(f"Invalid portal URL: {url}")
    return f"{parsed.scheme}://{parsed.netloc}"


class UnitePortalClient:
    def __init__(self, base_url: str, timeout: int = 25) -> None:
        self.base_url = normalize_url(base_url)
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        })
        self.is_logged_in = False

    def login(self, username: str, password: str, on_status: Callable[[str], None] | None = None) -> bool:
        def status(msg: str) -> None:
            if on_status:
                on_status(msg)

        status(f"Connecting to portal {self.base_url}...")
        try:
            r = self.session.get(self.base_url, timeout=self.timeout, verify=False)
        except Exception as exc:
            raise UnitePortalError(f"Could not connect to {self.base_url}: {exc}") from exc

        soup = BeautifulSoup(r.text, "html.parser")
        token_el = soup.find("input", {"name": "__RequestVerificationToken"})
        if not token_el or not token_el.get("value"):
            raise UnitePortalError(
                f"Verification token not found at {self.base_url}. Check if the portal is accessible."
            )
        token = token_el["value"]

        status("Submitting portal login credentials...")
        login_data = {
            "UserName": username.strip(),
            "Password": password,
            "command": "Login",
            "__RequestVerificationToken": token,
        }
        try:
            login_res = self.session.post(
                f"{self.base_url}/",
                data=login_data,
                timeout=self.timeout,
                verify=False,
                allow_redirects=True,
            )
        except Exception as exc:
            raise UnitePortalError(f"Login request failed: {exc}") from exc

        # Check login success: either redirected to menuscreen or cookies set
        resp_text = login_res.text.lower()
        if "menuscreen" in login_res.url.lower() or "btnlogout" in resp_text or "membership" in resp_text:
            self.is_logged_in = True
            status("Portal login successful!")
            return True

        if "invalid" in resp_text or "incorrect" in resp_text or "login" in login_res.url.lower():
            raise UnitePortalError("Login failed! Please check URL, Login ID, or Password.")

        # If cookies present, proceed
        if self.session.cookies:
            self.is_logged_in = True
            status("Portal login authenticated.")
            return True

        raise UnitePortalError("Login verification failed. Portal returned unexpected response.")

    def get_districts(self, state_id: int = STATE_UP_ID) -> list[dict]:
        url = f"{self.base_url}/Membership/CaptureCustomerLandDetails/GetDistrict"
        try:
            r = self.session.get(url, params={"intStateid": state_id}, timeout=self.timeout, verify=False)
            data = r.json().get("data", [])
            return data
        except Exception as exc:
            raise UnitePortalError(f"Failed to fetch districts: {exc}") from exc

    def get_mandals(self, district_pkey: int) -> list[dict]:
        url = f"{self.base_url}/Membership/CaptureCustomerLandDetails/GetMandal"
        try:
            r = self.session.get(url, params={"Distict": district_pkey}, timeout=self.timeout, verify=False)
            data = r.json().get("data", [])
            return data
        except Exception as exc:
            raise UnitePortalError(f"Failed to fetch mandals: {exc}") from exc

    def get_villages(self, mandal_pkey: int) -> list[str]:
        url = f"{self.base_url}/Membership/CaptureCustomerLandDetails/GetVillage"
        try:
            r = self.session.get(url, params={"Mandal": mandal_pkey}, timeout=self.timeout, verify=False)
            data = r.json().get("data", [])
        except Exception as exc:
            raise UnitePortalError(f"Failed to fetch villages: {exc}") from exc

        valid: list[str] = []
        seen: set[str] = set()
        for item in data:
            name = str(item.get("VillageName", "")).strip()
            if is_invalid_village(name):
                continue
            key = name.lower()
            if key not in seen:
                seen.add(key)
                valid.append(name)
        return valid

    def _fetch_one_mandal_villages(self, m: dict) -> tuple[str, list[str]]:
        """Fetch villages for a single mandal (used in parallel execution)."""
        m_name = str(m.get("MandalName", "")).strip()
        m_pkey = m.get("Pkey")
        if not m_name or not m_pkey:
            return m_name, []
        villages = self.get_villages(m_pkey)
        return m_name, villages

    def fetch_membership_pacs_villages(
        self,
        on_status: Callable[[str], None] | None = None,
    ) -> list[str]:
        """Fetch exact PACS-registered villages from active portal session:
        Membership -> Personal Details -> Capture Customer Details (+Add) -> Village* dropdown.
        Strictly excludes '--Select--', 'Others', 'Other', etc.
        Returns every legitimate village option present in the active PACS's portal dropdown.
        """
        def status(msg: str) -> None:
            if on_status:
                on_status(msg)

        status("Navigating to Membership -> Personal Details -> + Add (formid=1001)...")
        # Direct URL as observed in portal UI (+ Add button):
        urls_to_try = [
            f"{self.base_url}/Membership?formid=1001&moduleid=1",
            f"{self.base_url}/Membership/CaptureCustomerPersonalDetails/CaptureCustomerPersonalDetails?formid=1001&moduleid=1",
            f"{self.base_url}/Membership/CaptureCustomerPersonalDetails?formid=1001&moduleid=1",
            f"{self.base_url}/Membership/CaptureCustomerDetails?formid=1001&moduleid=1",
            f"{self.base_url}/Membership/CustomerPersonalDetails",
            f"{self.base_url}/Membership",
        ]

        resp = None
        for u in urls_to_try:
            try:
                r = self.session.get(u, timeout=10, verify=False)
                if r.status_code == 200 and ("village" in r.text.lower() or "customer" in r.text.lower()):
                    resp = r
                    break
            except Exception:
                continue

        if resp is None:
            status("Membership +Add page not accessible via HTTP GET, falling back to registered portal villages.")
            return []

        soup = BeautifulSoup(resp.text, "html.parser")
        select_el = None

        # 1. Select with id or name containing 'village' (Primary: VillageCode)
        for s in soup.find_all("select"):
            sid = f"{s.get('id', '')} {s.get('name', '')}".lower()
            if "village" in sid:
                select_el = s
                break

        # 2. Label matching for 'village'
        if not select_el:
            for lbl in soup.find_all(["label", "span", "div", "td", "th"]):
                if "village" in lbl.text.lower():
                    parent = lbl.find_parent(["div", "tr", "td", "section", "fieldset"])
                    if parent:
                        s = parent.find("select")
                        if s:
                            select_el = s
                            break

        # 3. Fallback: Select with 'others' option having more than 5 options (distinguishes from GenderCode)
        if not select_el:
            for s in soup.find_all("select"):
                opts = s.find_all("option")
                if len(opts) > 5:
                    opt_texts = [o.text.strip().lower() for o in opts]
                    if any(x in opt_texts for x in ["others", "other", "अन्य"]):
                        select_el = s
                        break

        if not select_el:
            status("PACS-specific dropdown not present in static HTML, falling back to registered portal villages.")
            return []

        valid_villages: list[str] = []
        for opt in select_el.find_all("option"):
            name = opt.text.strip()
            if is_invalid_village(name):
                continue
            if name not in valid_villages:
                valid_villages.append(name)

        if not valid_villages:
            status("PACS dropdown has no valid registered villages, falling back to registered portal villages.")
            return []

        status(f"[OK] Live PACS Membership villages loaded from portal dropdown ({len(valid_villages)} villages, strictly no 'Others'): {', '.join(valid_villages)}")
        return valid_villages

    def fetch_villages_for_district(
        self,
        district_query: str,
        on_status: Callable[[str], None] | None = None,
    ) -> tuple[list[str], LandPortalData]:
        """Fetch ALL mandals and their villages in PARALLEL for speed."""
        def status(msg: str) -> None:
            if on_status:
                on_status(msg)

        status("Fetching UP districts list from portal...")
        districts = self.get_districts()
        if not districts:
            raise UnitePortalError("No districts returned from portal.")

        # Find matching district: STRICT EXACT MATCH FIRST
        q = district_query.strip().lower()
        matched_dist = None

        # 1. Exact match (case-insensitive) - e.g. "RAMPUR" must match "RAMPUR", NOT "BALRAMPUR"
        for d in districts:
            d_name = str(d.get("DistrictName", "")).strip().lower()
            if q == d_name:
                matched_dist = d
                break

        # 2. Known UP District Aliases / Alternative Spellings
        if not matched_dist:
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
            alias_target = aliases.get(q)
            if alias_target:
                for d in districts:
                    d_name = str(d.get("DistrictName", "")).strip().lower()
                    if alias_target == d_name or alias_target in d_name:
                        matched_dist = d
                        break

        # 3. Word-boundary match
        if not matched_dist:
            for d in districts:
                d_name = str(d.get("DistrictName", "")).strip().lower()
                if q in d_name.split():
                    matched_dist = d
                    break

        # 4. Substring match
        if not matched_dist:
            for d in districts:
                d_name = str(d.get("DistrictName", "")).strip().lower()
                if q in d_name:
                    matched_dist = d
                    break

        # 5. Fuzzy match with difflib
        if not matched_dist:
            import difflib
            dist_names = [str(d.get("DistrictName", "")).strip().lower() for d in districts]
            close = difflib.get_close_matches(q, dist_names, n=1, cutoff=0.65)
            if close:
                best_match = close[0]
                for d in districts:
                    if str(d.get("DistrictName", "")).strip().lower() == best_match:
                        matched_dist = d
                        break

        if not matched_dist:
            avail = ", ".join(str(d.get("DistrictName", "")) for d in districts[:8])
            raise UnitePortalError(
                f"District '{district_query}' not found on portal. Available examples: {avail}..."
            )

        dist_name = matched_dist.get("DistrictName", district_query).strip()
        dist_pkey = matched_dist.get("Pkey")
        status(f"District matched: {dist_name} (ID: {dist_pkey})")

        status(f"Fetching all Mandals for {dist_name}...")
        mandals = self.get_mandals(dist_pkey)
        if not mandals:
            raise UnitePortalError(f"No mandals found for district '{dist_name}'.")

        # Process ALL available Mandals for the District as per portal Land Details mapping
        mandals_to_process = mandals
        status(f"Processing all {len(mandals)} available mandal(s) for District '{dist_name}'...")

        total = len(mandals_to_process)
        status(f"Fetching villages for {total} mandal(s) in parallel...")

        # --- PARALLEL fetch ---
        mandal_village_map: dict[str, list[str]] = {}
        all_villages: list[str] = []
        done_count = 0

        with ThreadPoolExecutor(max_workers=min(total, 8)) as executor:
            future_map = {
                executor.submit(self._fetch_one_mandal_villages, m): m
                for m in mandals_to_process
            }
            for future in as_completed(future_map):
                done_count += 1
                try:
                    m_name, villages = future.result()
                    if villages:
                        mandal_village_map[m_name] = villages
                        all_villages.extend(v for v in villages if v not in all_villages)
                        status(f"  [{done_count}/{total}] {m_name}: {len(villages)} villages")
                    else:
                        status(f"  [{done_count}/{total}] Skipped - no villages")
                except Exception as exc:
                    status(f"  [{done_count}/{total}] Error fetching mandal: {exc}")

        if not mandal_village_map:
            raise UnitePortalError(f"No mandals with valid villages for '{dist_name}'.")

        total_v = sum(len(v) for v in mandal_village_map.values())
        status(f"Portal OK - {len(mandal_village_map)} mandal(s), {total_v} total villages fetched.")

        land_data = LandPortalData(
            state_name="UTTAR PRADESH",
            district_desc=dist_name,
            mandal_village_map=mandal_village_map,
        )
        return all_villages, land_data

