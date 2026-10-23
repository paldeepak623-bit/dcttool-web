"""Portal Session Manager — fast direct HTTP connection with fallback."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

from dcttool.portals.unite_portal import (
    LandPortalData,
    UnitePortalClient,
    UnitePortalError,
    normalize_url,
)


@dataclass(frozen=True)
class PortalSessionData:
    villages: list[str] | None = None
    land_data: LandPortalData | None = None
    is_pacs_dropdown: bool = False


def fetch_portal_session(
    url: str,
    login_id: str,
    password: str,
    *,
    need_villages: bool,
    need_land: bool,
    district_desc: str = "",
    headless: bool = True,
    on_status: Callable[[str], None] | None = None,
) -> PortalSessionData:
    """One fast HTTP login for all portal data (Membership villages + Land data)."""
    if not need_villages and not need_land:
        return PortalSessionData()

    if not login_id.strip():
        raise UnitePortalError("Login ID is required for portal village/mandal mapping.")

    district = district_desc.strip()
    if not district:
        raise UnitePortalError("District is required for portal village mapping.")

    def status(msg: str) -> None:
        if on_status:
            on_status(msg)

    client = UnitePortalClient(url)
    try:
        client.login(login_id, password, on_status=status)
    except Exception as login_err:
        if need_villages:
            raise UnitePortalError(
                f"Portal Login Failed for '{login_id}' at {url}: {login_err}\n\n"
                f"Membership template me kewal wahi gaon bhare ja sakte hain jo is PACS ke portal dropdown me registered hain.\n"
                f"Bina PACS portal login ke genuine gaon prapt nahi ho sakte, jisse portal par sabhi members 'Others' ho jayenge.\n"
                f"Kripya sahi Portal URL, Login ID aur Password daalkar verify karein."
            ) from login_err
        status(f"Portal login note: {login_err}")

    villages = None
    land_data = None
    is_pacs_dropdown = False

    # 1. High-Speed Priority: Fetch exact PACS-registered villages directly from active portal dropdown
    # (Membership -> Personal Details -> + Add -> Village* dropdown)
    if need_villages:
        try:
            status("Checking live PACS-specific dropdown on portal (Membership -> Personal Details -> + Add)...")
            pacs_dropdown_villages = client.fetch_membership_pacs_villages(on_status=status)
            if pacs_dropdown_villages:
                villages = pacs_dropdown_villages
                is_pacs_dropdown = True
                status(f"[PORTAL] Successfully loaded {len(villages)} PACS-specific dropdown village(s): {', '.join(villages[:5])}...")
            else:
                raise UnitePortalError(
                    "PACS ke portal dropdown (Membership -> Personal Details -> + Add) me koi registered village nahi mila.\n"
                    "Portal par bina valid dropdown village ke data dalne par sabhi members 'Others' ho jate hain.\n"
                    "Kripya portal par verify karein ki is PACS ke liye gaon register hain ya nahi."
                )
        except Exception as e:
            raise UnitePortalError(f"PACS portal dropdown se gaon prapt karne me error: {e}") from e

    # 2. Only fetch district mandal/land mapping if Land Details is selected
    if need_land:
        status(f"Fetching official UP portal registered villages for District '{district}'...")
        _, land_data = client.fetch_villages_for_district(district, on_status=status)

    return PortalSessionData(
        villages=villages,
        land_data=land_data,
        is_pacs_dropdown=is_pacs_dropdown,
    )

