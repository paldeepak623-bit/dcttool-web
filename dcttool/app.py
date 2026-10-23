# -*- coding: utf-8 -*-
"""
DCTtool V2.0 Enterprise - Spectrum Pastel Edition
Matching reference designs:
- Dashboard (System dashboard & Cockpit Telemetry)
- Setup / Config (Configure your process & Pastel Template Library)
- Process Engine (Live process log & Telemetry Monitor)
- Admin Settings (Admin & automation settings & Brand Showcase)
Dual Theme: Light Mode (Spectrum Pastel) & Dark Mode with instant toggle switch.
100% of Business Rules & Template Processing Decoupled & Fully Operational.
Developed by Deepak Pal.
(C) 2026 Deepak Pal. All Rights Reserved.
"""

from __future__ import annotations

import datetime
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Callable

import customtkinter as ctk
from PIL import Image

from dcttool import __app_name__, __version__
from dcttool.master_engine import (
    MasterEngineError,
    MasterSettings,
    TemplateSelection,
    format_sheet_data_popup,
    process_master_workbook,
    save_split_templates,
)
from dcttool.template_files import copy_templates_to_folder, list_bundled_templates
from dcttool.sheet_populator import extract_pacs_metadata_from_workbook

# Explicit Windows App User Model ID
try:
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("dcttool.enterprise.v2.deepakpal")
except Exception:
    pass


class Palette:
    # APPLICATION BACKGROUNDS (Section 2 & 3 - Spectrum Pastel)
    BG_BASE = "#FAF5FD"
    BG_MINT = "#E6FFF7"
    BG_CYAN = "#E3F7FF"
    BG_LAVENDER = "#F0ECFF"
    BG_PINK = "#FCEBFF"

    # Main App Background
    BG = ("#FAF5FD", "#0B1120")

    # SURFACES
    SURFACE_CARD = ("#FFFFFF", "#11223B")
    SURFACE_SUBTLE = ("#FDFBFE", "#0F1E33")
    SURFACE_INPUT = ("#FFFFFF", "#0B1626")
    SURFACE_CONSOLE = ("#EFF7FF", "#060C16")

    # Left Navigation Sidebar (Slim vertical nav, clean frosted surface)
    SIDEBAR_BG = ("#FCF8FD", "#070E1A")
    SIDEBAR_BORDER = ("#E2E8F0", "#1E293B")
    SIDEBAR_HOVER = ("#F5EFF8", "#111C30")
    SIDEBAR_INACTIVE_TXT = ("#52618A", "#F1F5F9")
    SIDEBAR_ACTIVE_BG = ("#2454E6", "#2563EB")
    SIDEBAR_ACTIVE_TXT = "#FFFFFF"

    # Top Header & Cockpit Telemetry Bar (Section 6 & 8)
    HEADER_BG = ("#FFFFFF", "#0D182A")
    HEADER_BORDER = ("#DCE5F3", "#1E293B")
    COCKPIT_BG = ("#FFFFFF", "#0F1E33")
    COCKPIT_BORDER = ("#DCE5F3", "#1E2E48")

    # Main Panels / Cards
    PANEL_BG = ("#FFFFFF", "#11223B")
    PANEL_BORDER = ("#DCE5F3", "#1E3252")

    # Input Fields & Comboboxes (Section 11)
    INPUT_BG = ("#FFFFFF", "#0B1626")
    INPUT_BORDER = ("#C8D5EA", "#213C60")
    BORDER_DEFAULT = ("#DCE5F3", "#1E3252")
    BORDER_DIVIDER = ("#E4EAF5", "#1E2E48")
    BORDER_FOCUS = "#2454E6"

    # Dropzone (Master sheet input box - Section 13)
    DROPZONE_BG = ("#F0FCFF", "#0D251E")
    DROPZONE_BORDER = ("#22B8CF", "#00838F")

    # Terminal / Console Log Box (Section 15 & 16)
    CONSOLE_BG = ("#EFF7FF", "#060C16")
    CONSOLE_BORDER = ("#D7E7F8", "#15283F")
    CONSOLE_TEXT = ("#18234B", "#4ADE80")

    # Typography (Section 2 & 4)
    TEXT_HEADING = ("#101044", "#F8FAFC")
    TEXT_TITLE = ("#101044", "#F8FAFC")
    TEXT_PRIMARY = ("#18234B", "#E2E8F0")
    TEXT_MAIN = ("#18234B", "#E2E8F0")
    TEXT_SECONDARY = ("#52618A", "#94A3B8")
    TEXT_MUTED = ("#52618A", "#94A3B8")
    TEXT_PLACEHOLDER = ("#697692", "#64748B")
    TEXT_DISABLED = ("#78839A", "#64748B")

    # Brand Colors (Section 2)
    BRAND_BLUE = "#2454E6"
    BRAND_VIOLET = "#6D28D9"
    BRAND_PURPLE = "#8527DE"
    BRAND_CYAN = "#06B6D4"

    TEXT_BLUE = ("#2454E6", "#38BDF8")
    TEXT_GREEN = ("#087A50", "#4ADE80")
    TEXT_AMBER = ("#B45309", "#FBBF24")
    TEXT_PURPLE = ("#6D28D9", "#A78BFA")
    TEXT_WHITE = "#FFFFFF"

    # Buttons (Section 10)
    BTN_GRADIENT_BG = "#2454E6"
    BTN_GRADIENT_HOVER = "#1D45C8"
    BTN_SECONDARY_BG = ("#FFFFFF", "#132742")
    BTN_SECONDARY_BORDER = ("#2454E6", "#224168")
    BTN_SECONDARY_TXT = ("#2454E6", "#38BDF8")
    BTN_SECONDARY_HOVER = ("#EDF3FF", "#1E3A64")

    # Status Badges (Section 2)
    BADGE_READY_BG = ("#DDF9EC", "#064E3B")
    BADGE_READY_TXT = ("#087A50", "#34D399")
    BADGE_OFF_BG = ("#F1F5F9", "#1E293B")
    BADGE_OFF_TXT = ("#78839A", "#94A3B8")

    # Template Library Specific Colors (Section 14)
    # 1. Declaration: Fill #DDF8FF, Accent/icon #087E9A
    DECL_BG = ("#DDF8FF", "#0C2B44")
    DECL_BORDER = ("#2454E6", "#0284C7")
    DECL_TXT = ("#087E9A", "#7DD3FC")

    # 2. Short Term Loans: Fill #DFFBF0, Accent/icon #087A50
    ST_BG = ("#DFFBF0", "#063A24")
    ST_BORDER = ("#2454E6", "#16A34A")
    ST_TXT = ("#087A50", "#86EFAC")

    # 3. Membership: Fill #EEE5FF, Accent/icon #6D28D9
    MEM_BG = ("#EEE5FF", "#241846")
    MEM_BORDER = ("#2454E6", "#7C3AED")
    MEM_TXT = ("#6D28D9", "#C4B5FD")

    # 4. Land Details: Fill #FCE3F4, Accent/icon #A21C85
    LAND_BG = ("#FCE3F4", "#3D1027")
    LAND_BORDER = ("#2454E6", "#DB2777")
    LAND_TXT = ("#A21C85", "#F9A8D4")

    # 5. Savings: Fill #FFEADB, Accent/icon #B45309
    SAV_BG = ("#FFEADB", "#3A230A")
    SAV_BORDER = ("#2454E6", "#D97706")
    SAV_TXT = ("#B45309", "#FDE68A")

    # 6. Loan Transaction: Fill #FFF7D6, Accent/icon #896000
    LT_BG = ("#FFF7D6", "#3A2E05")
    LT_BORDER = ("#2454E6", "#D97706")
    LT_TXT = ("#896000", "#FDE68A")

    # 7. Dropdown List: Fill #F1F5FA, Accent/icon #52618A
    DD_BG = ("#F1F5FA", "#1E293B")
    DD_BORDER = ("#DCE5F3", "#334155")
    DD_TXT = ("#52618A", "#94A3B8")

    # 8. Loan Conversion: Fill #E8EEFF, Accent/icon #2454E6
    LC_BG = ("#E8EEFF", "#172554")
    LC_BORDER = ("#DCE5F3", "#3B82F6")
    LC_TXT = ("#2454E6", "#93C5FD")

    # Backward compatibility aliases
    GRAY_BG = DD_BG
    GRAY_BORDER = DD_BORDER
    GRAY_TXT = DD_TXT


# Primary Font: Segoe UI, Fallback: Arial (Section 4)
FONT_FAMILY = "Segoe UI"
FONT_MONO = "Cascadia Mono"


def font_main(size: int = 11, weight: str = "normal", slant: str = "roman") -> ctk.CTkFont:
    """Standardized Segoe UI font constructor across the entire application."""
    return ctk.CTkFont(family=FONT_FAMILY, size=size, weight=weight, slant=slant)


def font_mono(size: int = 11, weight: str = "normal") -> ctk.CTkFont:
    """Standardized Cascadia Mono font constructor for logs and telemetry."""
    return ctk.CTkFont(family=FONT_MONO, size=size, weight=weight)


# Maintain font_tnr as alias so all existing calls resolve cleanly to Segoe UI
font_tnr = font_main

UP_DISTRICTS = [
    "AGRA", "ALIGARH", "AMBEDKAR NAGAR", "AMETHI", "AMROHA", "AURAIYA", "AYODHYA", "AZAMGARH",
    "BAGHPAT", "BAHRAICH", "BALLIA", "BALRAMPUR", "BANDA", "BARABANKI", "BAREILLY", "BASTI",
    "BHADOHI", "BIJNOR", "BUDAUN", "BULANDSHAHR", "CHANDAULI", "CHITRAKOOT", "DEORIA", "ETAH",
    "ETAWAH", "FARRUKHABAD", "FATEHPUR", "FIROZABAD", "GAUTAM BUDDHA NAGAR", "GHAZIABAD", "GHAZIPUR",
    "GONDA", "GORAKHPUR", "HAMIRPUR", "HAPUR", "HARDOI", "HATHRAS", "JALAUN", "JAUNPUR",
    "JHANSI", "KANNAUJ", "KANPUR DEHAT", "KANPUR NAGAR", "KASGANJ", "KAUSHAMBI", "KHERI",
    "KUSHINAGAR", "LALITPUR", "LUCKNOW", "MAHARAJGANJ", "MAHOBA", "MAINPURI", "MATHURA", "MAU",
    "MEERUT", "MIRZAPUR", "MORADABAD", "MUZAFFARNAGAR", "PILIBHIT", "PRATAPGARH", "PRAYAGRAJ",
    "RAE BARELI", "RAMPUR", "SAHARANPUR", "SAMBHAL", "SANT KABIR NAGAR", "SHAHJAHANPUR", "SHAMLI",
    "SHRAVASTI", "SIDDHARTHNAGAR", "SITAPUR", "SONBHADRA", "SULTANPUR", "UNNAO", "VARANASI"
]

PORTAL_URLS = [
    "https://up3.unitetools.in/",
    "https://up1.unitetools.in/",
    "https://up2.unitetools.in/",
    "https://up4.unitetools.in/",
    "https://up5.unitetools.in/",
    "https://up6.unitetools.in/"
]

CONFIG_DIR = Path.home() / ".dcttool"
CONFIG_FILE = CONFIG_DIR / "config.json"


def load_persisted_config() -> dict[str, str]:
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_persisted_config(data: dict[str, str]) -> None:
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass



def bind_all_children(widget, sequence: str, func: Callable):
    """Recursively bind an event to a widget and all its children."""
    try:
        widget.bind(sequence, func)
    except Exception:
        pass
    if hasattr(widget, "_canvas"):
        try:
            widget._canvas.bind(sequence, func)
        except Exception:
            pass
    for child in widget.winfo_children():
        bind_all_children(child, sequence, func)


class VisualAssets:
    _cache: dict[str, ctk.CTkImage] = {}

    PRIMARY_STOPS = [(0.0, (36, 84, 230)), (0.55, (109, 40, 217)), (1.0, (133, 39, 222))]
    PRIMARY_HOVER_STOPS = [(0.0, (29, 69, 200)), (0.55, (91, 33, 182)), (1.0, (114, 32, 195))]
    SIDEBAR_STOPS = [(0.0, (36, 84, 230)), (1.0, (109, 40, 217))]

    @classmethod
    def get_rounded_gradient(
        cls, width: int, height: int, radius: int, stops: list, key: str
    ) -> ctk.CTkImage:
        cache_key = f"{key}_{width}_{height}_{radius}"
        if cache_key in cls._cache:
            return cls._cache[cache_key]

        scale = 2
        sw, sh, sr = max(width * scale, 4), max(height * scale, 4), max(radius * scale, 1)
        row = Image.new("RGBA", (sw, 1))
        for x in range(sw):
            t = x / max(sw - 1, 1)
            for i in range(len(stops) - 1):
                t0, c0 = stops[i]
                t1, c1 = stops[i + 1]
                if t0 <= t <= t1 or i == len(stops) - 2:
                    factor = (t - t0) / max(t1 - t0, 0.0001)
                    r = int(c0[0] + (c1[0] - c0[0]) * factor)
                    g = int(c0[1] + (c1[1] - c0[1]) * factor)
                    b = int(c0[2] + (c1[2] - c0[2]) * factor)
                    row.putpixel((x, 0), (r, g, b, 255))
                    break
        grad = row.resize((sw, sh), Image.Resampling.NEAREST)
        from PIL import ImageDraw
        mask = Image.new("L", (sw, sh), 0)
        draw = ImageDraw.Draw(mask)
        draw.rounded_rectangle((0, 0, sw - 1, sh - 1), radius=sr, fill=255)
        result = Image.new("RGBA", (sw, sh), (0, 0, 0, 0))
        result.paste(grad, (0, 0), mask=mask)
        final_img = result.resize((width, height), Image.Resampling.LANCZOS)
        ctk_img = ctk.CTkImage(light_image=final_img, dark_image=final_img, size=(width, height))
        cls._cache[cache_key] = ctk_img
        return ctk_img

    @classmethod
    def create_gradient_button(
        cls,
        parent,
        text: str,
        command: Callable,
        width: int = 220,
        height: int = 44,
        font: ctk.CTkFont | None = None,
        corner_radius: int = 10,
    ) -> ctk.CTkButton:
        btn = ctk.CTkButton(
            parent,
            text=text,
            width=width,
            height=height,
            corner_radius=corner_radius,
            fg_color="#2454E6",
            hover_color="#1D45C8",
            text_color="#FFFFFF",
            font=font or font_main(12, weight="bold"),
            command=command,
        )
        return btn


class DCTtoolApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        # Default to Light appearance mode matching Spectrum Pastel screenshots
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title("DCTtool - v2.0.0 Enterprise - PACS Master Automation Platform")
        self.geometry("1360x780")
        self.minsize(1050, 600)
        try:
            self.after(50, lambda: self.state("zoomed"))
        except Exception:
            pass
        self.configure(fg_color=Palette.BG)

        # Assets Directory
        meipass = getattr(sys, "_MEIPASS", None)
        self.base_dir = Path(meipass) if meipass else Path(__file__).resolve().parent.parent
        self.assets_dir = self.base_dir / "assets"

        # Load Visual Graphic Assets
        self._load_assets()

        # Core State Variables
        self.sheet_path: Path | None = None
        self._processed_wb = None
        self._process_report = None
        self._log_queue: queue.Queue[str] = queue.Queue()
        self._worker: threading.Thread | None = None
        self._show_pwd = False
        self.active_view = "setup"

        # Reactive Variables (Loaded from persisted config or current active PACS defaults)
        cfg = load_persisted_config()
        self.pacs_var = ctk.StringVar(value=cfg.get("pacs_id", "UP320052618339"))
        self.district_var = ctk.StringVar(value=cfg.get("district", "LAKHIMPUR KHERI"))
        self.url_var = ctk.StringVar(value=cfg.get("url", "https://up3.unitetools.in/"))
        self.login_var = ctk.StringVar(value=cfg.get("login_id", "UP320052618339_de@coopsindia.com"))
        self.password_var = ctk.StringVar(value=cfg.get("password", "Unite@123"))

        self.file_title_var = ctk.StringVar(value="B-PACS RAJAPURKALA (PYAGPUR).xlsx")
        self.file_sub_var = ctk.StringVar(value="File loaded successfully\nC:\\Users\\Deepak Pal\\Downloads")
        self.file_loaded_flag = False

        self.progress_pct_var = ctk.StringVar(value="100%")
        self.progress_role_var = ctk.StringVar(value="● Ready")
        self.target_pacs_label_var = ctk.StringVar(value=f"[Cite: 1] ({self.district_var.get().title()})")
        self.engine_status_var = ctk.StringVar(value="Online (100% Load)")
        self.theme_status_var = ctk.StringVar(value="Light")

        # District, PACS ID and Login synchronization
        self.district_var.trace_add("write", self._on_district_change)
        self.pacs_var.trace_add("write", self._on_pacs_change)
        self.login_var.trace_add("write", self._on_login_change)


        # Template Selection Checkbox Variables
        self.chk_declaration = ctk.BooleanVar(value=True)
        self.chk_stloans = ctk.BooleanVar(value=True)
        self.chk_membership = ctk.BooleanVar(value=True)
        self.chk_land = ctk.BooleanVar(value=True)
        self.chk_savings = ctk.BooleanVar(value=True)
        self.chk_loan = ctk.BooleanVar(value=True)
        self.chk_dropdown = ctk.BooleanVar(value=False)
        self.chk_stloans_to_lt = ctk.BooleanVar(value=True)

        self._sidebar_buttons: dict[str, ctk.CTkButton] = {}
        self._template_cards: dict[str, dict] = {}
        self._views: dict[str, ctk.CTkFrame] = {}

        # Build Interface
        self._build_ui()
        self.after(100, self._drain_log)

    def _on_district_change(self, *args) -> None:
        d = self.district_var.get().strip().upper()
        self.target_pacs_label_var.set(f"[Cite: 1] ({d.title()})")
        self._log(f"[CONFIG] Target District set to: {d}")

    def _on_pacs_change(self, *args) -> None:
        p = self.pacs_var.get().strip()
        self._log(f"[CONFIG] Target PACS ID set to: {p}")

    def _on_login_change(self, *args) -> None:
        raw = self.login_var.get().strip().upper()
        for i in range(1, 10):
            if raw.startswith(f"UP{i}"):
                matched_url = f"https://up{i}.unitetools.in/"
                if self.url_var.get() != matched_url:
                    self.url_var.set(matched_url)
                    self._log(f"[PORTAL AUTO] Auto-aligned Portal URL to: {matched_url} from Login ID")
                break


    def _load_assets(self) -> None:
        def _get_img(filename: str, size: tuple[int, int]) -> ctk.CTkImage | None:
            p = self.assets_dir / filename
            if p.exists():
                try:
                    pil_im = Image.open(p)
                    return ctk.CTkImage(light_image=pil_im, dark_image=pil_im, size=size)
                except Exception:
                    return None
            return None

        self.img_icon = _get_img("icon.png", (28, 28))
        self.img_map = _get_img("map_preview.png", (140, 70))

        # High-definition Navigation Icons
        self.nav_icons = {}
        for k in ("dashboard", "setup", "process", "data_hub", "settings"):
            self.nav_icons[k] = {
                "active": _get_img(f"nav_{k}_active.png", (20, 20)),
                "inactive": _get_img(f"nav_{k}_inactive.png", (20, 20))
            }

        ico_path = self.assets_dir / "icon.ico"
        if ico_path.exists():
            try:
                self.iconbitmap(default=str(ico_path))
            except Exception:
                pass

        self._views_built: dict[str, bool] = {}

    def _build_ui(self) -> None:
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=0, minsize=96)
        self.grid_columnconfigure(1, weight=1)

        self._build_sidebar()
        self._build_body()

    def _build_sidebar(self) -> None:
        sidebar = ctk.CTkFrame(
            self, width=96, corner_radius=0,
            fg_color=Palette.SIDEBAR_BG,
            border_color=Palette.SIDEBAR_BORDER, border_width=1
        )
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        # Top App Logo Container
        logo_box = ctk.CTkFrame(sidebar, fg_color="transparent")
        logo_box.pack(pady=(16, 12))

        # Vibrant Blue App Icon Badge
        badge = ctk.CTkButton(
            logo_box, text="D1", width=48, height=48,
            corner_radius=12, fg_color="#2454E6", hover_color="#1D45C8",
            text_color="#FFFFFF", font=font_main(18, weight="bold"),
            command=lambda: self._switch_view("dashboard")
        )
        badge.pack()

        # Nav Items Stack (Vertical icon above text layout matching Spectrum Pastel reference)
        navs = [
            ("Dashboard", "dashboard"),
            ("Setup / Config", "setup"),
            ("Process Engine", "process"),
            ("Data Hub", "data_hub"),
            ("Admin Settings", "settings"),
        ]

        for title, key in navs:
            is_active = (key == "setup")
            ico = self.nav_icons.get(key, {}).get("active" if is_active else "inactive")
            btn = ctk.CTkButton(
                sidebar,
                text=title,
                image=ico,
                compound="top",
                width=84,
                height=58,
                corner_radius=12,
                fg_color="#2454E6" if is_active else "transparent",
                text_color="#FFFFFF" if is_active else "#52618A",
                hover_color="#1D45C8" if is_active else "#F1F5F9",
                font=font_main(9, weight="bold" if is_active else "normal"),
                command=lambda k=key: self._switch_view(k)
            )
            btn.pack(pady=4, padx=5)
            self._sidebar_buttons[key] = btn

        # Bottom Developer Attribution
        sidebar_foot = ctk.CTkFrame(sidebar, fg_color="transparent")
        sidebar_foot.pack(side="bottom", pady=12, padx=4)

        ctk.CTkLabel(
            sidebar_foot,
            text=f"DCTtool v{__version__}\nEnterprise\n\n© 2026 Deepak Pal.\nAll Rights Reserved.",
            font=font_main(8),
            text_color=Palette.TEXT_MUTED,
            justify="center"
        ).pack()

    def _init_view(self, key: str) -> None:
        """Lazily initialize a view only when navigated to or during idle."""
        if self._views_built.get(key, False):
            return
        frame = self._views.get(key)
        if not frame:
            return
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        if key == "setup":
            self._build_setup_view(frame)
        elif key == "dashboard":
            self._build_dashboard_view(frame)
        elif key == "process":
            self._build_process_engine_view(frame)
        elif key == "data_hub":
            self._build_data_hub_view(frame)
        elif key == "settings":
            self._build_admin_settings_view(frame)
        self._views_built[key] = True

    def _prebuild_other_views(self) -> None:
        """Prebuild remaining views in background idle so tab switching is instantaneous."""
        for k in ("dashboard", "process", "data_hub", "settings"):
            if not self._views_built.get(k, False):
                self._init_view(k)

    def _switch_view(self, key: str) -> None:
        """Switch views dynamically between Dashboard, Setup, Process, Data Hub and Settings."""
        if key not in self._views:
            return
        if not self._views_built.get(key, False):
            self._init_view(key)

        self.active_view = key

        # Update sidebar active styling
        for btn_key, btn in self._sidebar_buttons.items():
            is_active = (btn_key == key)
            ico = self.nav_icons.get(btn_key, {}).get("active" if is_active else "inactive")
            if is_active:
                btn.configure(
                    image=ico,
                    compound="top",
                    fg_color="#2454E6",
                    hover_color="#1D45C8",
                    text_color="#FFFFFF",
                    font=font_main(9, weight="bold")
                )
            else:
                btn.configure(
                    image=ico,
                    compound="top",
                    fg_color="transparent",
                    hover_color="#F1F5F9",
                    text_color="#52618A",
                    font=font_main(9, weight="normal")
                )

        # Raise target view frame
        self._views[key].tkraise()
        self._update_workflow_stepper(key)
        self._log(f"[NAV] Switched view to: {key.replace('_', ' ').title()}")

    def _update_workflow_stepper(self, view_key: str) -> None:
        if hasattr(self, "step_pill1") and hasattr(self, "step_pill2") and hasattr(self, "step_pill3"):
            if view_key in ("setup", "dashboard"):
                self.step_pill1.configure(fg_color="#6D28D9", text_color="#FFFFFF")
                self.step_txt1.configure(text_color="#2454E6", font=font_tnr(11, weight="bold"))
                self.step_pill2.configure(fg_color="#E5EAF3", text_color="#52618A")
                self.step_txt2.configure(text_color="#52618A", font=font_tnr(11))
                self.step_pill3.configure(fg_color="#E5EAF3", text_color="#52618A")
                self.step_txt3.configure(text_color="#52618A", font=font_tnr(11))
            elif view_key == "process":
                self.step_pill1.configure(fg_color="#E5EAF3", text_color="#52618A")
                self.step_txt1.configure(text_color="#52618A", font=font_tnr(11))
                self.step_pill2.configure(fg_color="#E5EAF3", text_color="#52618A")
                self.step_txt2.configure(text_color="#52618A", font=font_tnr(11))
                self.step_pill3.configure(fg_color="#6D28D9", text_color="#FFFFFF")
                self.step_txt3.configure(text_color="#2454E6", font=font_tnr(11, weight="bold"))

    def _build_body(self) -> None:
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=0, column=1, sticky="nsew", padx=14, pady=10)
        body.grid_rowconfigure(2, weight=1)
        body.grid_columnconfigure(0, weight=1)

        # 1. Top Header Bar
        self._build_top_header(body)

        # 2. Cockpit Telemetry Cards Row
        self._build_cockpit_row(body)

        # 3. View Container (Holds all interactive views)
        view_container = ctk.CTkFrame(body, fg_color="transparent")
        view_container.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        view_container.grid_rowconfigure(0, weight=1)
        view_container.grid_columnconfigure(0, weight=1)

        # View 1: Setup / Config (Main Cockpit Workspace - built immediately)
        v_setup = ctk.CTkFrame(view_container, fg_color="transparent")
        v_setup.grid(row=0, column=0, sticky="nsew")
        v_setup.grid_rowconfigure(0, weight=1)
        v_setup.grid_columnconfigure(0, weight=1)
        self._views["setup"] = v_setup
        self._init_view("setup")

        # Views 2-5: Registered for on-demand lazy initialization
        for k in ("dashboard", "process", "data_hub", "settings"):
            vf = ctk.CTkFrame(view_container, fg_color="transparent")
            vf.grid(row=0, column=0, sticky="nsew")
            self._views[k] = vf
            self._views_built[k] = False

        # Raise setup view by default
        self._views["setup"].tkraise()
        # Schedule idle background prebuilding of remaining views
        self.after(300, self._prebuild_other_views)

    def _build_top_header(self, parent: ctk.CTkFrame) -> None:
        header = ctk.CTkFrame(
            parent, fg_color=Palette.HEADER_BG, corner_radius=12,
            border_color=Palette.HEADER_BORDER, border_width=1, height=52
        )
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header.pack_propagate(False)

        h_inner = ctk.CTkFrame(header, fg_color="transparent")
        h_inner.pack(fill="both", expand=True, padx=16, pady=4)

        # Left Branding
        left_box = ctk.CTkFrame(h_inner, fg_color="transparent")
        left_box.pack(side="left")

        # Blue badge icon with "D1"
        ctk.CTkLabel(
            left_box, text="D1", width=32, height=32,
            fg_color="#2454E6", text_color="#FFFFFF",
            corner_radius=8, font=font_tnr(14, weight="bold")
        ).pack(side="left", padx=(0, 8))

        ctk.CTkLabel(
            left_box, text="DCTtool",
            font=font_tnr(19, weight="bold"),
            text_color=Palette.TEXT_HEADING
        ).pack(side="left")

        ctk.CTkLabel(
            left_box, text="  |   v2.0.0 Enterprise - PACS Master Automation Platform",
            font=font_tnr(11),
            text_color=Palette.TEXT_SECONDARY
        ).pack(side="left", padx=4)

        # Right Controls: Account Pill, Online Badge, Theme Toggle
        right_box = ctk.CTkFrame(h_inner, fg_color="transparent")
        right_box.pack(side="right")

        # Account Pill
        acct_btn = ctk.CTkButton(
            right_box,
            text="👤  Account: Deepak Pal  |  ID: DP0012345678",
            font=font_tnr(11, weight="bold"),
            fg_color=("#F8FAFF", "#0F1E33"), hover_color=Palette.SIDEBAR_HOVER,
            border_color=("#DCE5F3", "#1E2E48"), border_width=1,
            text_color=("#18234B", "#E2E8F0"), corner_radius=8, height=32,
            command=lambda: self._switch_view("settings")
        )
        acct_btn.pack(side="left", padx=(0, 10))

        # Online Status Pill
        ctk.CTkLabel(
            right_box, text="●  Online",
            font=font_tnr(12, weight="bold"),
            text_color="#087A50"
        ).pack(side="left", padx=(0, 12))

        # Theme Toggle: Sun ☀️, Switch, Moon 🌙
        ctk.CTkLabel(
            right_box, text="☀️", font=font_tnr(13), text_color=Palette.TEXT_MUTED
        ).pack(side="left", padx=(0, 4))

        self.theme_switch = ctk.CTkSwitch(
            right_box, text="", variable=self.theme_status_var,
            onvalue="Dark", offvalue="Light",
            width=42, height=22, switch_width=18, switch_height=18,
            progress_color="#6366F1",
            command=self._toggle_theme
        )
        # Default is Light
        self.theme_switch.deselect()
        self.theme_switch.pack(side="left", padx=(0, 4))

        ctk.CTkLabel(
            right_box, text="🌙", font=font_tnr(12), text_color=Palette.TEXT_MUTED
        ).pack(side="left", padx=(0, 10))

        # Window Control Buttons
        for sym in ("—", "▢", "✕"):
            cmd = None
            if sym == "—":
                cmd = self.iconify
            elif sym == "▢":
                cmd = self._toggle_maximize
            elif sym == "✕":
                cmd = self.destroy
            ctk.CTkButton(
                right_box, text=sym, width=28, height=28, corner_radius=6,
                fg_color="transparent", hover_color=Palette.SIDEBAR_HOVER,
                text_color=Palette.TEXT_MUTED, font=font_tnr(11, weight="bold"),
                command=cmd
            ).pack(side="left", padx=1)

    def _toggle_maximize(self) -> None:
        if self.state() == "zoomed":
            self.state("normal")
        else:
            self.state("zoomed")

    def _toggle_theme(self) -> None:
        mode = self.theme_switch.get()
        if mode == "Dark":
            ctk.set_appearance_mode("dark")
            self._log("[THEME] Appearance switched to: Dark Mode")
        else:
            ctk.set_appearance_mode("light")
            self._log("[THEME] Appearance switched to: Light Mode (Spectrum Pastel)")

    def _build_cockpit_row(self, parent: ctk.CTkFrame) -> None:
        """Top telemetry capsule row matching all screenshots."""
        row = ctk.CTkFrame(
            parent, fg_color=Palette.COCKPIT_BG, corner_radius=12,
            border_color=Palette.COCKPIT_BORDER, border_width=1, height=64
        )
        row.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        row.pack_propagate(False)

        r_inner = ctk.CTkFrame(row, fg_color="transparent")
        r_inner.pack(fill="both", expand=True, padx=14, pady=6)
        for i in range(5):
            r_inner.grid_columnconfigure(i, weight=1)

        # Block 1: Engine Status
        b1 = ctk.CTkFrame(r_inner, fg_color="transparent", cursor="hand2")
        b1.grid(row=0, column=0, sticky="nsew", padx=6)
        ctk.CTkLabel(b1, text="🟢  Engine Status", font=font_tnr(10), text_color=Palette.TEXT_MUTED).pack(anchor="w")
        ctk.CTkLabel(b1, textvariable=self.engine_status_var, font=font_tnr(13, weight="bold"), text_color=Palette.TEXT_GREEN).pack(anchor="w")
        bind_all_children(b1, "<Button-1>", lambda e: self._switch_view("process"))

        # Block 2: Target PACS
        b2 = ctk.CTkFrame(r_inner, fg_color="transparent", cursor="hand2")
        b2.grid(row=0, column=1, sticky="nsew", padx=6)
        ctk.CTkLabel(b2, text="🔲  Target PACS", font=font_tnr(10), text_color=Palette.TEXT_MUTED).pack(anchor="w")
        ctk.CTkLabel(b2, textvariable=self.target_pacs_label_var, font=font_tnr(13, weight="bold"), text_color=Palette.TEXT_TITLE).pack(anchor="w")
        bind_all_children(b2, "<Button-1>", lambda e: self._focus_pacs_input())

        # Block 3: Select District ComboBox
        b3 = ctk.CTkFrame(r_inner, fg_color="transparent")
        b3.grid(row=0, column=2, sticky="ew", padx=8)
        ctk.CTkLabel(b3, text="📍  Select District", font=font_tnr(10), text_color=Palette.TEXT_MUTED).pack(anchor="w")
        self.dist_combo = ctk.CTkComboBox(
            b3, values=UP_DISTRICTS, variable=self.district_var,
            height=28, font=font_tnr(11, weight="bold"),
            fg_color=Palette.INPUT_BG, border_color=Palette.INPUT_BORDER,
            button_color="#2563EB", text_color=Palette.TEXT_TITLE,
            dropdown_fg_color=Palette.PANEL_BG,
            command=lambda val: self.district_var.set(val)
        )
        self.dist_combo.pack(fill="x", pady=(2, 0))

        # Block 4: Processed Declarations
        b4 = ctk.CTkFrame(r_inner, fg_color="transparent", cursor="hand2")
        b4.grid(row=0, column=3, sticky="nsew", padx=6)
        ctk.CTkLabel(b4, text="📄  Processed Declarations", font=font_tnr(10), text_color=Palette.TEXT_MUTED).pack(anchor="w")
        ctk.CTkLabel(b4, text="5.2k", font=font_tnr(14, weight="bold"), text_color=Palette.TEXT_TITLE).pack(anchor="w")
        bind_all_children(b4, "<Button-1>", lambda e: self._switch_view("process"))

        # Block 5: Processed Loans
        b5 = ctk.CTkFrame(r_inner, fg_color="transparent", cursor="hand2")
        b5.grid(row=0, column=4, sticky="nsew", padx=6)
        ctk.CTkLabel(b5, text="📄  Processed Loans", font=font_tnr(10), text_color=Palette.TEXT_MUTED).pack(anchor="w")
        ctk.CTkLabel(b5, text="3.1k", font=font_tnr(14, weight="bold"), text_color=Palette.TEXT_TITLE).pack(anchor="w")
        bind_all_children(b5, "<Button-1>", lambda e: self._switch_view("process"))

    def _focus_pacs_input(self) -> None:
        self._switch_view("setup")
        if hasattr(self, "login_entry"):
            self.login_entry.focus()

    def _build_page_title_bar(self, parent: ctk.CTkFrame, title: str, subtitle: str) -> None:
        """Top section of each page containing title, subtitle and 3-step workflow breadcrumbs."""
        title_f = ctk.CTkFrame(parent, fg_color="transparent")
        title_f.pack(fill="x", pady=(0, 10))

        # Left Title & Subtitle
        left_t = ctk.CTkFrame(title_f, fg_color="transparent")
        left_t.pack(side="left")

        ctk.CTkLabel(
            left_t, text=title,
            font=font_tnr(22, weight="bold"),
            text_color=Palette.TEXT_TITLE
        ).pack(anchor="w")

        ctk.CTkLabel(
            left_t, text=subtitle,
            font=font_tnr(11),
            text_color=Palette.TEXT_MUTED
        ).pack(anchor="w", pady=(2, 0))

        # Right Stepper Pill
        stepper = ctk.CTkFrame(title_f, fg_color="transparent")
        stepper.pack(side="right", pady=4)

        # Step 1
        self.step_pill1 = ctk.CTkLabel(
            stepper, text="1", width=22, height=22, corner_radius=11,
            fg_color="#6D28D9", text_color="#FFFFFF",
            font=font_tnr(11, weight="bold")
        )
        self.step_pill1.pack(side="left", padx=(0, 6))

        self.step_txt1 = ctk.CTkLabel(
            stepper, text="Data input",
            font=font_tnr(11, weight="bold"),
            text_color="#2454E6"
        )
        self.step_txt1.pack(side="left", padx=(0, 12))

        ctk.CTkLabel(stepper, text="──", font=font_tnr(10), text_color="#CBD5E1").pack(side="left", padx=(0, 12))

        # Step 2
        self.step_pill2 = ctk.CTkLabel(
            stepper, text="2", width=22, height=22, corner_radius=11,
            fg_color="#E5EAF3", text_color="#52618A",
            font=font_tnr(11, weight="bold")
        )
        self.step_pill2.pack(side="left", padx=(0, 6))

        self.step_txt2 = ctk.CTkLabel(
            stepper, text="Templates",
            font=font_tnr(11),
            text_color="#52618A"
        )
        self.step_txt2.pack(side="left", padx=(0, 12))

        ctk.CTkLabel(stepper, text="──", font=font_tnr(10), text_color="#CBD5E1").pack(side="left", padx=(0, 12))

        # Step 3
        self.step_pill3 = ctk.CTkLabel(
            stepper, text="3", width=22, height=22, corner_radius=11,
            fg_color="#E5EAF3", text_color="#52618A",
            font=font_tnr(11, weight="bold")
        )
        self.step_pill3.pack(side="left", padx=(0, 6))

        self.step_txt3 = ctk.CTkLabel(
            stepper, text="Execute / Monitor",
            font=font_tnr(11),
            text_color="#52618A"
        )
        self.step_txt3.pack(side="left")

    # -------------------------------------------------------------------------
    # VIEW 1: SETUP / CONFIG (Image 4 - Spectrum Pastel)
    # -------------------------------------------------------------------------
    def _build_setup_view(self, parent: ctk.CTkFrame) -> None:
        self._build_page_title_bar(
            parent,
            title="Configure your process",
            subtitle="Complete the steps below to input data, configure templates and execute the automation process."
        )

        # ---------------------------------------------------------------------
        # 1. ALWAYS DOCK ACTION BAR AT THE VERY BOTTOM OF THE WINDOW (NEVER CLIPS!)
        # ---------------------------------------------------------------------
        action_bar = ctk.CTkFrame(
            parent, fg_color=Palette.PANEL_BG, corner_radius=12,
            border_color=Palette.PANEL_BORDER, border_width=1, height=52
        )
        action_bar.pack(side="bottom", fill="x", pady=(6, 2))
        action_bar.pack_propagate(False)

        ab_inner = ctk.CTkFrame(action_bar, fg_color="transparent")
        ab_inner.pack(fill="both", expand=True, padx=14, pady=6)

        # Left File Pill
        left_pill = ctk.CTkFrame(ab_inner, fg_color="transparent")
        left_pill.pack(side="left")

        ctk.CTkLabel(
            left_pill, text="📄", font=font_tnr(13)
        ).pack(side="left", padx=(0, 6))

        ctk.CTkLabel(
            left_pill, text="Master sheet: ",
            font=font_tnr(11, weight="bold"),
            text_color=Palette.TEXT_TITLE
        ).pack(side="left")

        self.lbl_master_pill = ctk.CTkLabel(
            left_pill, textvariable=self.file_title_var,
            font=font_tnr(11, weight="bold"),
            text_color=Palette.TEXT_BLUE
        )
        self.lbl_master_pill.pack(side="left")

        ctk.CTkLabel(
            left_pill, text="  ✔  |  1 file loaded  |  6 templates selected",
            font=font_tnr(10),
            text_color=Palette.TEXT_MUTED
        ).pack(side="left", padx=4)

        # Right Radiant Start Process Button
        self.btn_start = VisualAssets.create_gradient_button(
            ab_inner, text="Start Process",
            command=self._start_process_thread,
            width=200, height=38, corner_radius=10,
            font=font_main(12, weight="bold")
        )
        self.btn_start.pack(side="right")

        # ---------------------------------------------------------------------
        # 2. MAIN WORKSPACE CONTENT (Fills between title bar and action bar)
        # ---------------------------------------------------------------------
        content = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        content.pack(side="top", fill="both", expand=True)

        # 3 Top Cards Row (Master sheet, Portal config, Execution status)
        top_row = ctk.CTkFrame(content, fg_color="transparent")
        top_row.pack(fill="x", pady=(0, 6))
        top_row.grid_rowconfigure(0, weight=1)
        top_row.grid_columnconfigure(0, weight=4)
        top_row.grid_columnconfigure(1, weight=5)
        top_row.grid_columnconfigure(2, weight=4)

        # 1. Card 1: Master sheet data input
        p1 = ctk.CTkFrame(
            top_row, fg_color=Palette.PANEL_BG, corner_radius=12,
            border_color=Palette.PANEL_BORDER, border_width=1, height=285
        )
        p1.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        h1 = ctk.CTkFrame(p1, fg_color="transparent")
        h1.pack(fill="x", padx=12, pady=(10, 2))
        ctk.CTkLabel(h1, text="📄", font=font_tnr(13)).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(h1, text="Master sheet data input", font=font_tnr(12, weight="bold"), text_color=Palette.TEXT_TITLE).pack(side="left")
        ctk.CTkLabel(p1, text="Upload your PACS master sheet in Excel format.", font=font_tnr(9), text_color=Palette.TEXT_MUTED).pack(anchor="w", padx=12, pady=(0, 6))

        # Dropzone Container
        self.drop_card = ctk.CTkFrame(
            p1, fg_color=Palette.DROPZONE_BG, corner_radius=10,
            border_color=Palette.DROPZONE_BORDER, border_width=1, height=130, cursor="hand2"
        )
        self.drop_card.pack(fill="x", padx=12, pady=(0, 8))
        self.drop_card.pack_propagate(False)

        drop_inner = ctk.CTkFrame(self.drop_card, fg_color="transparent")
        drop_inner.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            drop_inner, text="📊 ✔", font=font_tnr(16, weight="bold"), text_color="#10B981"
        ).pack(pady=(0, 2))

        self.lbl_file_title = ctk.CTkLabel(
            drop_inner, textvariable=self.file_title_var,
            font=font_tnr(11, weight="bold"),
            text_color=Palette.TEXT_TITLE
        )
        self.lbl_file_title.pack()

        self.lbl_file_sub = ctk.CTkLabel(
            drop_inner, textvariable=self.file_sub_var,
            font=font_tnr(9),
            text_color=Palette.TEXT_MUTED, justify="center"
        )
        self.lbl_file_sub.pack(pady=(2, 0))

        bind_all_children(self.drop_card, "<Button-1>", lambda e: self._browse_file())

        # Buttons Row: Replace file | Get sample templates | Remove
        btn_box = ctk.CTkFrame(p1, fg_color="transparent")
        btn_box.pack(fill="x", padx=12, pady=(0, 6))

        ctk.CTkButton(
            btn_box, text="🔄 Replace file", height=28, corner_radius=8,
            fg_color=Palette.BTN_SECONDARY_BG, hover_color=Palette.BTN_SECONDARY_HOVER,
            border_color=Palette.BTN_SECONDARY_BORDER, border_width=1,
            text_color=Palette.TEXT_BLUE, font=font_tnr(10, weight="bold"),
            command=self._browse_file
        ).pack(side="left", fill="x", expand=True, padx=(0, 2))

        ctk.CTkButton(
            btn_box, text="📥 Sample templates", height=28, corner_radius=8,
            fg_color="transparent", hover_color=Palette.SIDEBAR_HOVER,
            border_color=Palette.BTN_SECONDARY_BORDER, border_width=1,
            text_color=Palette.TEXT_BLUE, font=font_tnr(10, weight="bold"),
            command=self._download_sample_templates
        ).pack(side="left", fill="x", expand=True, padx=(2, 2))

        self.btn_remove_file = ctk.CTkButton(
            btn_box, text="✕ Remove", height=28, corner_radius=8,
            fg_color="#FEE2E2", hover_color="#FECACA",
            border_color="#FCA5A5", border_width=1,
            text_color="#DC2626", font=font_tnr(10, weight="bold"),
            command=self._remove_file
        )

        # 2. Card 2: Portal configuration
        p2 = ctk.CTkFrame(
            top_row, fg_color=Palette.PANEL_BG, corner_radius=12,
            border_color=Palette.PANEL_BORDER, border_width=1, height=285
        )
        p2.grid(row=0, column=1, sticky="nsew", padx=3)

        h2 = ctk.CTkFrame(p2, fg_color="transparent")
        h2.pack(fill="x", padx=12, pady=(10, 2))
        ctk.CTkLabel(h2, text="🔗", font=font_tnr(13)).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(h2, text="Portal configuration", font=font_tnr(12, weight="bold"), text_color=Palette.TEXT_TITLE).pack(side="left")
        ctk.CTkLabel(p2, text="Enter portal credentials for data processing.", font=font_tnr(9), text_color=Palette.TEXT_MUTED).pack(anchor="w", padx=12, pady=(0, 6))

        p2_body = ctk.CTkFrame(p2, fg_color="transparent")
        p2_body.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        p2_body.grid_columnconfigure(0, weight=3)
        p2_body.grid_columnconfigure(1, weight=2)

        form = ctk.CTkFrame(p2_body, fg_color="transparent")
        form.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        ctk.CTkLabel(form, text="Portal URL", font=font_main(10, weight="bold"), text_color=Palette.TEXT_MUTED).pack(anchor="w")
        self.url_combo = ctk.CTkComboBox(
            form, values=PORTAL_URLS, variable=self.url_var, height=28,
            font=font_main(11), fg_color=Palette.INPUT_BG,
            border_color=Palette.INPUT_BORDER, button_color="#2454E6", text_color=Palette.TEXT_TITLE
        )
        self.url_combo.pack(fill="x", pady=(1, 4))

        ctk.CTkLabel(form, text="Login ID", font=font_main(10, weight="bold"), text_color=Palette.TEXT_MUTED).pack(anchor="w")
        self.login_entry = ctk.CTkEntry(
            form, textvariable=self.login_var, height=28,
            font=font_main(11), fg_color=Palette.INPUT_BG,
            border_color=Palette.INPUT_BORDER, text_color=Palette.TEXT_TITLE
        )
        self.login_entry.pack(fill="x", pady=(1, 4))

        ctk.CTkLabel(form, text="Password", font=font_main(10, weight="bold"), text_color=Palette.TEXT_MUTED).pack(anchor="w")
        pwd_box = ctk.CTkFrame(form, fg_color="transparent")
        pwd_box.pack(fill="x", pady=(1, 0))

        self.pwd_entry = ctk.CTkEntry(
            pwd_box, textvariable=self.password_var, show="*", height=28,
            font=font_main(11), fg_color=Palette.INPUT_BG,
            border_color=Palette.INPUT_BORDER, text_color=Palette.TEXT_TITLE
        )
        self.pwd_entry.pack(side="left", fill="x", expand=True)

        self.btn_eye = ctk.CTkButton(
            pwd_box, text="👁️", width=36, height=28, corner_radius=6,
            fg_color=Palette.INPUT_BG, border_color=Palette.INPUT_BORDER, border_width=1,
            hover_color=Palette.SIDEBAR_HOVER, font=font_main(11),
            command=self._toggle_password_view
        )
        self.btn_eye.pack(side="left", padx=(4, 0))

        # Mini District Map Preview Card
        map_box = ctk.CTkFrame(
            p2_body, fg_color=Palette.INPUT_BG, corner_radius=10,
            border_color=Palette.INPUT_BORDER, border_width=1
        )
        map_box.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(
            map_box, text="District map preview",
            font=font_main(9), text_color=Palette.TEXT_MUTED
        ).pack(pady=(4, 0))

        if self.img_map:
            ctk.CTkLabel(map_box, image=self.img_map, text="").pack(expand=True, pady=1)
        else:
            ctk.CTkLabel(map_box, text="🗺️ Map View", font=font_main(11), text_color=Palette.TEXT_BLUE).pack(expand=True)

        ctk.CTkLabel(
            map_box, text="District: BAHRAICH",
            font=font_main(9, weight="bold"), text_color=Palette.TEXT_BLUE
        ).pack(pady=(0, 4))

        # 3. Card 3: Execution status & Live Log
        p3 = ctk.CTkFrame(
            top_row, fg_color=Palette.PANEL_BG, corner_radius=12,
            border_color=Palette.PANEL_BORDER, border_width=1, height=285
        )
        p3.grid(row=0, column=2, sticky="nsew", padx=(6, 0))

        h3 = ctk.CTkFrame(p3, fg_color="transparent")
        h3.pack(fill="x", padx=12, pady=(10, 2))
        ctk.CTkLabel(h3, text="▶", font=font_tnr(12), text_color=Palette.TEXT_GREEN).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(h3, text="Execution status", font=font_tnr(12, weight="bold"), text_color=Palette.TEXT_TITLE).pack(side="left")

        # Ready badge pill
        badge_ready = ctk.CTkLabel(
            h3, textvariable=self.progress_role_var,
            font=font_tnr(9, weight="bold"),
            fg_color=Palette.BADGE_READY_BG, text_color=Palette.BADGE_READY_TXT,
            corner_radius=12, height=20, padx=8
        )
        badge_ready.pack(side="right")

        ctk.CTkLabel(
            p3, text="System is ready to process. All configurations are valid.",
            font=font_tnr(9), text_color=Palette.TEXT_MUTED
        ).pack(anchor="w", padx=12, pady=(0, 4))

        # Overall progress bar
        prog_top = ctk.CTkFrame(p3, fg_color="transparent")
        prog_top.pack(fill="x", padx=12)
        ctk.CTkLabel(prog_top, text="Overall progress", font=font_tnr(9, weight="bold"), text_color=Palette.TEXT_MUTED).pack(side="left")
        ctk.CTkLabel(prog_top, textvariable=self.progress_pct_var, font=font_tnr(10, weight="bold"), text_color=Palette.TEXT_TITLE).pack(side="right")

        self.progress_bar = ctk.CTkProgressBar(p3, height=6, corner_radius=3, progress_color="#10B981")
        self.progress_bar.pack(fill="x", padx=12, pady=(2, 6))
        self.progress_bar.set(1.0)

        # 4 Checklist items
        checks = [
            ("Master sheet loaded", "10:14 AM"),
            ("Portal connection verified", "10:14 AM"),
            ("Templates configured", "10:15 AM"),
            ("System ready for execution", "10:15 AM"),
        ]
        chk_box = ctk.CTkFrame(p3, fg_color="transparent")
        chk_box.pack(fill="x", padx=12, pady=(0, 4))
        for item, tm in checks:
            row_c = ctk.CTkFrame(chk_box, fg_color="transparent")
            row_c.pack(fill="x", pady=1)
            ctk.CTkLabel(row_c, text="✔", font=font_tnr(9, weight="bold"), text_color="#10B981").pack(side="left", padx=(0, 6))
            ctk.CTkLabel(row_c, text=item, font=font_tnr(9), text_color=Palette.TEXT_MAIN).pack(side="left")
            ctk.CTkLabel(row_c, text=tm, font=font_tnr(9), text_color=Palette.TEXT_MUTED).pack(side="right")

        # ---------------------------------------------------------------------
        # Bottom Split: Template Library (Left) & Live Process Log (Right)
        # ---------------------------------------------------------------------
        bottom_split = ctk.CTkFrame(content, fg_color="transparent")
        bottom_split.pack(fill="both", expand=True, pady=(0, 4))
        bottom_split.grid_columnconfigure(0, weight=6)
        bottom_split.grid_columnconfigure(1, weight=4)

        # Left: Template library
        lib_frame = ctk.CTkFrame(
            bottom_split, fg_color=Palette.PANEL_BG, corner_radius=12,
            border_color=Palette.PANEL_BORDER, border_width=1
        )
        lib_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        lib_head = ctk.CTkFrame(lib_frame, fg_color="transparent")
        lib_head.pack(fill="x", padx=14, pady=(8, 2))

        ctk.CTkLabel(lib_head, text="📋", font=font_tnr(13)).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(
            lib_head, text="Template library",
            font=font_tnr(12, weight="bold"),
            text_color=Palette.TEXT_TITLE
        ).pack(side="left")

        # Select all / Clear buttons on right
        btn_clear = ctk.CTkButton(
            lib_head, text="✕ Clear selection", width=95, height=22,
            fg_color="transparent", hover_color=Palette.SIDEBAR_HOVER,
            text_color=Palette.TEXT_MUTED, font=font_tnr(9),
            command=self._select_none
        )
        btn_clear.pack(side="right", padx=(4, 0))

        btn_sel_all = ctk.CTkButton(
            lib_head, text="☑ Select all (7/8)", width=105, height=22,
            fg_color="transparent", hover_color=Palette.SIDEBAR_HOVER,
            text_color=Palette.TEXT_BLUE, font=font_tnr(9, weight="bold"),
            command=self._select_all
        )
        btn_sel_all.pack(side="right")

        ctk.CTkLabel(
            lib_frame, text="Select the templates to process. You can choose multiple templates based on your requirement.",
            font=font_tnr(9), text_color=Palette.TEXT_MUTED
        ).pack(anchor="w", padx=14, pady=(0, 4))

        # 8 Pastel Cards Grid (2 rows x 4 cols)
        grid_f = ctk.CTkFrame(lib_frame, fg_color="transparent")
        grid_f.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        for c_idx in range(4):
            grid_f.grid_columnconfigure(c_idx, weight=1)

        templates_def = [
            ("declaration", "Declaration", "📄", Palette.DECL_BG, Palette.DECL_BORDER, Palette.DECL_TXT, self.chk_declaration, 0, 0),
            ("stloans", "Short Term Loans", "₹", Palette.ST_BG, Palette.ST_BORDER, Palette.ST_TXT, self.chk_stloans, 0, 1),
            ("membership", "Membership", "👥", Palette.MEM_BG, Palette.MEM_BORDER, Palette.MEM_TXT, self.chk_membership, 0, 2),
            ("land", "Land Details", "🎯", Palette.LAND_BG, Palette.LAND_BORDER, Palette.LAND_TXT, self.chk_land, 0, 3),
            ("savings", "Savings", "🏛️", Palette.SAV_BG, Palette.SAV_BORDER, Palette.SAV_TXT, self.chk_savings, 1, 0),
            ("loan", "Loan Transaction", "💳", Palette.LT_BG, Palette.LT_BORDER, Palette.LT_TXT, self.chk_loan, 1, 1),
            ("dropdown", "Dropdown List", "📑", Palette.DD_BG, Palette.DD_BORDER, Palette.DD_TXT, self.chk_dropdown, 1, 2),
            ("stloans_to_lt", "ShortTermLoans To LT", "🔄", Palette.LC_BG, Palette.LC_BORDER, Palette.LC_TXT, self.chk_stloans_to_lt, 1, 3),
        ]

        for key, name, icon, bg_col, b_col, t_col, var, r, c in templates_def:
            card = self._create_pastel_tile(grid_f, key, name, icon, bg_col, b_col, t_col, var, r, c)
            self._template_cards[key] = card

        # Right: Live process log card
        log_card = ctk.CTkFrame(
            bottom_split, fg_color=Palette.PANEL_BG, corner_radius=12,
            border_color=Palette.PANEL_BORDER, border_width=1
        )
        log_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        log_head = ctk.CTkFrame(log_card, fg_color="transparent")
        log_head.pack(fill="x", padx=12, pady=(8, 2))

        ctk.CTkLabel(log_head, text="📋", font=font_tnr(13)).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(
            log_head, text="Live process log",
            font=font_tnr(12, weight="bold"),
            text_color=Palette.TEXT_TITLE
        ).pack(side="left")

        ctk.CTkButton(
            log_head, text="📥 Download logs", height=22, corner_radius=6,
            fg_color=Palette.INPUT_BG, border_color=Palette.INPUT_BORDER, border_width=1,
            text_color=Palette.TEXT_BLUE, font=font_tnr(9, weight="bold"),
            command=self._save_logs
        ).pack(side="right")

        ctk.CTkButton(
            log_head, text="🗑️ Clear logs", height=22, corner_radius=6,
            fg_color=Palette.INPUT_BG, border_color=Palette.INPUT_BORDER, border_width=1,
            text_color=Palette.TEXT_MUTED, font=font_tnr(9, weight="bold"),
            command=self._clear_logs
        ).pack(side="right", padx=(0, 6))

        ctk.CTkLabel(
            log_card, text="Real-time execution logs and process status.",
            font=font_tnr(9), text_color=Palette.TEXT_MUTED
        ).pack(anchor="w", padx=12, pady=(0, 4))

        # Log Text Box
        self.log_textbox = ctk.CTkTextbox(
            log_card, font=font_mono(10),
            fg_color=Palette.CONSOLE_BG, border_color=Palette.CONSOLE_BORDER,
            border_width=1, text_color=Palette.CONSOLE_TEXT, corner_radius=8
        )
        self.log_textbox.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        self._seed_default_logs()

    def _create_pastel_tile(
        self, parent: ctk.CTkFrame, key: str, name: str, icon: str,
        bg_col: tuple[str, str], b_col: tuple[str, str], t_col: tuple[str, str],
        var: ctk.BooleanVar, r: int, c: int
    ) -> dict:
        """Create an aesthetic pastel card matching Section 14 (selected vs unselected)."""
        tile = ctk.CTkFrame(
            parent, fg_color=bg_col, corner_radius=10,
            border_color=b_col, border_width=1, height=44, cursor="hand2"
        )
        tile.grid(row=r, column=c, padx=4, pady=3, sticky="nsew")
        tile.pack_propagate(False)

        t_inner = ctk.CTkFrame(tile, fg_color="transparent")
        t_inner.place(relx=0.5, rely=0.5, anchor="center")

        # Icon + Name
        ctk.CTkLabel(
            t_inner, text=icon, font=font_tnr(12), text_color=t_col
        ).pack(side="left", padx=(0, 5))

        ctk.CTkLabel(
            t_inner, text=name,
            font=font_tnr(10, weight="bold"),
            text_color=t_col
        ).pack(side="left", padx=(0, 6))

        # Checkbox
        chk = ctk.CTkCheckBox(
            t_inner, text="", variable=var, width=16, height=16,
            corner_radius=4, fg_color="#2454E6", hover_color="#1D45C8",
            border_color=b_col
        )
        chk.pack(side="right")

        def _update_tile_state(*args):
            try:
                if var.get():
                    tile.configure(fg_color=bg_col, border_color=b_col)
                else:
                    tile.configure(fg_color=("#FAFCFF", "#11223B"), border_color=("#DCE5F3", "#1E3252"))
            except Exception:
                pass

        var.trace_add("write", _update_tile_state)
        _update_tile_state()

        def _toggle(e=None):
            new_val = not var.get()
            var.set(new_val)
            self._log(f"[TEMPLATE] {name} is now: {'SELECTED' if new_val else 'UNCHECKED'}")

        bind_all_children(tile, "<Button-1>", _toggle)
        return {"frame": tile, "chk": chk, "var": var}

    def _seed_default_logs(self) -> None:
        samples = [
            "🟢 10:16:23  Successfully loaded master sheet: B-PACS RAJAPURKALA (PYAGPUR).xlsx",
            "🟢 10:16:25  Portal connection established with UP portal",
            "🟢 10:16:27  Templates validated (7 selected)",
            "🟢 10:16:30  System ready for process execution"
        ]
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.insert("end", "\n".join(samples) + "\n")
        self.log_textbox.see("end")

    def _select_all(self) -> None:
        for card in self._template_cards.values():
            card["var"].set(True)
        self._log("[TEMPLATE] Selected all templates.")

    def _select_none(self) -> None:
        for card in self._template_cards.values():
            card["var"].set(False)
        self._log("[TEMPLATE] Cleared template selection.")

    def _toggle_password_view(self) -> None:
        self._show_pwd = not self._show_pwd
        self.pwd_entry.configure(show="" if self._show_pwd else "*")

    # -------------------------------------------------------------------------
    # VIEW 2: SYSTEM DASHBOARD (Image 1 - Spectrum Pastel)
    # -------------------------------------------------------------------------
    def _build_dashboard_view(self, parent: ctk.CTkFrame) -> None:
        self._build_page_title_bar(
            parent,
            title="System dashboard",
            subtitle="Engine performance and quick actions."
        )

        content = ctk.CTkFrame(parent, fg_color="transparent")
        content.pack(fill="both", expand=True)

        card = ctk.CTkFrame(
            content, fg_color=Palette.PANEL_BG, corner_radius=14,
            border_color=Palette.PANEL_BORDER, border_width=1
        )
        card.pack(fill="both", expand=True)

        # Header
        ch = ctk.CTkFrame(card, fg_color="transparent")
        ch.pack(fill="x", padx=20, pady=(18, 6))

        ctk.CTkLabel(ch, text="📊", font=font_tnr(18)).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(
            ch, text="System Dashboard & Cockpit Telemetry",
            font=font_tnr(16, weight="bold"),
            text_color=Palette.TEXT_TITLE
        ).pack(side="left")

        ctk.CTkLabel(
            card, text="Overview of engine performance, loaded modules, and rapid operations.",
            font=font_tnr(11), text_color=Palette.TEXT_MUTED
        ).pack(anchor="w", padx=20, pady=(0, 16))

        # 3 Highlight Cards Row (Pastel Mint, Pastel Lavender, Pastel Peach)
        tri_row = ctk.CTkFrame(card, fg_color="transparent")
        tri_row.pack(fill="x", padx=20, pady=(0, 20))
        for i in range(3):
            tri_row.grid_columnconfigure(i, weight=1)

        # Card 1: Engine Status (Mint)
        c1 = ctk.CTkFrame(
            tri_row, fg_color=("#E5FBF3", "#063A24"), corner_radius=12,
            border_color=("#BAEEDC", "#16A34A"), border_width=1, height=92
        )
        c1.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        c1.pack_propagate(False)

        c1_h = ctk.CTkFrame(c1, fg_color="transparent")
        c1_h.pack(fill="both", expand=True, padx=16, pady=16)

        c1_tile = ctk.CTkFrame(c1_h, width=48, height=48, corner_radius=10, fg_color=("#C6F5E4", "#0A4D32"))
        c1_tile.pack(side="left", padx=(0, 14))
        c1_tile.pack_propagate(False)
        ctk.CTkLabel(c1_tile, text="⚡", font=font_main(20), text_color="#087A50").place(relx=0.5, rely=0.5, anchor="center")

        c1_txt = ctk.CTkFrame(c1_h, fg_color="transparent")
        c1_txt.pack(side="left", fill="both", expand=True)
        ctk.CTkLabel(c1_txt, text="Engine Status", font=font_main(11, weight="bold"), text_color="#52618A").pack(anchor="w")
        ctk.CTkLabel(c1_txt, text="ONLINE (100% Load)", font=font_main(16, weight="bold"), text_color="#087A50").pack(anchor="w", pady=(2, 0))

        # Card 2: Active District (Lavender)
        c2 = ctk.CTkFrame(
            tri_row, fg_color=("#F1EBFF", "#241846"), corner_radius=12,
            border_color=("#E0D4FA", "#7C3AED"), border_width=1, height=92
        )
        c2.grid(row=0, column=1, padx=5, sticky="nsew")
        c2.pack_propagate(False)

        c2_h = ctk.CTkFrame(c2, fg_color="transparent")
        c2_h.pack(fill="both", expand=True, padx=16, pady=16)

        c2_tile = ctk.CTkFrame(c2_h, width=48, height=48, corner_radius=10, fg_color=("#E3D6FF", "#331E62"))
        c2_tile.pack(side="left", padx=(0, 14))
        c2_tile.pack_propagate(False)
        ctk.CTkLabel(c2_tile, text="📍", font=font_main(20), text_color="#6D28D9").place(relx=0.5, rely=0.5, anchor="center")

        c2_txt = ctk.CTkFrame(c2_h, fg_color="transparent")
        c2_txt.pack(side="left", fill="both", expand=True)
        ctk.CTkLabel(c2_txt, text="Active District", font=font_main(11, weight="bold"), text_color="#52618A").pack(anchor="w")
        ctk.CTkLabel(c2_txt, textvariable=self.district_var, font=font_main(16, weight="bold"), text_color=Palette.TEXT_HEADING).pack(anchor="w", pady=(2, 0))

        # Card 3: Portal Gateway (Peach)
        c3 = ctk.CTkFrame(
            tri_row, fg_color=("#FFF1E6", "#3A230A"), corner_radius=12,
            border_color=("#F5DEC9", "#D97706"), border_width=1, height=92
        )
        c3.grid(row=0, column=2, padx=(10, 0), sticky="nsew")
        c3.pack_propagate(False)

        c3_h = ctk.CTkFrame(c3, fg_color="transparent")
        c3_h.pack(fill="both", expand=True, padx=16, pady=16)

        c3_tile = ctk.CTkFrame(c3_h, width=48, height=48, corner_radius=10, fg_color=("#FFE0C2", "#4D2C0C"))
        c3_tile.pack(side="left", padx=(0, 14))
        c3_tile.pack_propagate(False)
        ctk.CTkLabel(c3_tile, text="🌐", font=font_main(20), text_color="#B45309").place(relx=0.5, rely=0.5, anchor="center")

        c3_txt = ctk.CTkFrame(c3_h, fg_color="transparent")
        c3_txt.pack(side="left", fill="both", expand=True)
        ctk.CTkLabel(c3_txt, text="Portal Gateway", font=font_main(11, weight="bold"), text_color="#52618A").pack(anchor="w")
        ctk.CTkLabel(c3_txt, textvariable=self.url_var, font=font_main(14, weight="bold"), text_color=Palette.TEXT_HEADING).pack(anchor="w", pady=(2, 0))

        # Quick Actions Section
        qa_head = ctk.CTkFrame(card, fg_color="transparent")
        qa_head.pack(fill="x", padx=20, pady=(10, 4))
        ctk.CTkLabel(qa_head, text="⚡", font=font_main(14)).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(qa_head, text="Quick actions", font=font_main(13, weight="bold"), text_color=Palette.TEXT_TITLE).pack(side="left")

        ctk.CTkLabel(
            card, text="Navigate to key modules for configuration and monitoring.",
            font=font_main(10), text_color=Palette.TEXT_MUTED
        ).pack(anchor="w", padx=20, pady=(0, 14))

        qa_btns = ctk.CTkFrame(card, fg_color="transparent")
        qa_btns.pack(fill="x", padx=20, pady=(0, 20))

        VisualAssets.create_gradient_button(
            qa_btns, text="⚙️  Go to configuration & process  ➔",
            command=lambda: self._switch_view("setup"),
            width=270, height=44, font=font_main(12, weight="bold"), corner_radius=10
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            qa_btns, text="🗄️  Open Data Hub", height=44, corner_radius=10,
            fg_color=Palette.BTN_SECONDARY_BG, hover_color=Palette.BTN_SECONDARY_HOVER,
            border_color=Palette.BTN_SECONDARY_BORDER, border_width=1,
            text_color=Palette.BTN_SECONDARY_TXT, font=font_main(12, weight="bold"),
            command=lambda: self._switch_view("data_hub")
        ).pack(side="left", padx=10)

        ctk.CTkButton(
            qa_btns, text="📄  View real-time logs", height=44, corner_radius=10,
            fg_color=Palette.BTN_SECONDARY_BG, hover_color=Palette.BTN_SECONDARY_HOVER,
            border_color=Palette.BTN_SECONDARY_BORDER, border_width=1,
            text_color=Palette.BTN_SECONDARY_TXT, font=font_main(12, weight="bold"),
            command=lambda: self._switch_view("process")
        ).pack(side="left", padx=10)

    # -------------------------------------------------------------------------
    # VIEW 3: PROCESS ENGINE (Image 3 - Spectrum Pastel)
    # -------------------------------------------------------------------------
    def _build_process_engine_view(self, parent: ctk.CTkFrame) -> None:
        self._build_page_title_bar(
            parent,
            title="Process engine",
            subtitle="Real-time telemetry monitor."
        )

        content = ctk.CTkFrame(parent, fg_color="transparent")
        content.pack(fill="both", expand=True)

        card = ctk.CTkFrame(
            content, fg_color=Palette.PANEL_BG, corner_radius=14,
            border_color=Palette.PANEL_BORDER, border_width=1
        )
        card.pack(fill="both", expand=True)

        ch = ctk.CTkFrame(card, fg_color="transparent")
        ch.pack(fill="x", padx=20, pady=(16, 8))

        ctk.CTkLabel(ch, text=">_", font=font_tnr(14, weight="bold"), text_color=Palette.TEXT_BLUE).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(
            ch, text="Live process log",
            font=font_tnr(15, weight="bold"),
            text_color=Palette.TEXT_TITLE
        ).pack(side="left")

        ctk.CTkButton(
            ch, text="📥 Export full logs", height=28, corner_radius=8,
            fg_color=Palette.BTN_SECONDARY_BG, hover_color=Palette.BTN_SECONDARY_HOVER,
            border_color=Palette.BTN_SECONDARY_BORDER, border_width=1,
            text_color=Palette.TEXT_BLUE, font=font_tnr(10, weight="bold"),
            command=self._save_logs
        ).pack(side="right")

        # Console Textbox
        self.full_log_textbox = ctk.CTkTextbox(
            card, font=font_mono(10),
            fg_color=Palette.CONSOLE_BG, border_color=Palette.CONSOLE_BORDER,
            border_width=1, text_color=Palette.CONSOLE_TEXT, corner_radius=10
        )
        self.full_log_textbox.pack(fill="both", expand=True, padx=20, pady=(0, 16))
        self.full_log_textbox.insert("end", "\n\n\n\n               [ >_ ]\n            No logs yet\n   Start the master process to view execution logs.\n")

        # Bottom Action Buttons
        bot_dock = ctk.CTkFrame(card, fg_color="transparent")
        bot_dock.pack(fill="x", padx=20, pady=(0, 20))

        VisualAssets.create_gradient_button(
            bot_dock, text="  ▶  Start master process  ",
            command=self._start_process_thread,
            width=240, height=44, corner_radius=10,
            font=font_main(12, weight="bold")
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            bot_dock, text="🧹 Clear log console", height=40, corner_radius=10,
            fg_color=Palette.BTN_SECONDARY_BG, hover_color=Palette.BTN_SECONDARY_HOVER,
            border_color=Palette.BTN_SECONDARY_BORDER, border_width=1,
            text_color=Palette.BTN_SECONDARY_TXT, font=font_tnr(12, weight="bold"),
            command=self._clear_consoles
        ).pack(side="left")

    # -------------------------------------------------------------------------
    # VIEW 4: DATA HUB
    # -------------------------------------------------------------------------
    def _build_data_hub_view(self, parent: ctk.CTkFrame) -> None:
        self._build_page_title_bar(
            parent,
            title="Data Hub",
            subtitle="Explore generated templates and master sheets."
        )

        content = ctk.CTkFrame(parent, fg_color="transparent")
        content.pack(fill="both", expand=True)

        card = ctk.CTkFrame(
            content, fg_color=Palette.PANEL_BG, corner_radius=14,
            border_color=Palette.PANEL_BORDER, border_width=1
        )
        card.pack(fill="both", expand=True)

        ch = ctk.CTkFrame(card, fg_color="transparent")
        ch.pack(fill="x", padx=20, pady=(16, 8))

        ctk.CTkLabel(ch, text="🗄️", font=font_tnr(16)).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(
            ch, text="Data Hub - File & Output Explorer",
            font=font_tnr(15, weight="bold"),
            text_color=Palette.TEXT_TITLE
        ).pack(side="left")

        # Action Buttons
        btn_box = ctk.CTkFrame(card, fg_color="transparent")
        btn_box.pack(fill="x", padx=20, pady=(0, 12))

        ctk.CTkButton(
            btn_box, text="📁 Open Output Folder (DCT -1)", height=36, corner_radius=8,
            fg_color="#2454E6", hover_color="#1D45C8", text_color="#FFFFFF",
            font=font_tnr(11, weight="bold"),
            command=self._open_output_folder
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_box, text="📥 Export Bundled Templates to Desktop", height=36, corner_radius=8,
            fg_color=Palette.BTN_SECONDARY_BG, hover_color=Palette.BTN_SECONDARY_HOVER,
            border_color=Palette.BTN_SECONDARY_BORDER, border_width=1,
            text_color=Palette.BTN_SECONDARY_TXT, font=font_tnr(11, weight="bold"),
            command=self._export_templates_to_desktop
        ).pack(side="left")

        self.hub_listbox = ctk.CTkTextbox(
            card, font=font_mono(10),
            fg_color=Palette.CONSOLE_BG, border_color=Palette.CONSOLE_BORDER,
            border_width=1, text_color=Palette.CONSOLE_TEXT, corner_radius=10
        )
        self.hub_listbox.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self._refresh_data_hub_list()

    def _open_output_folder(self) -> None:
        target_dir = Path.home() / "Desktop" / "DCT -1"
        target_dir.mkdir(parents=True, exist_ok=True)
        try:
            subprocess.Popen(["explorer.exe", str(target_dir)])
            self._log(f"[HUB] Opened folder: {target_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder: {e}")

    def _export_templates_to_desktop(self) -> None:
        dest = Path.home() / "Desktop" / "DCT_Sample_Templates"
        dest.mkdir(parents=True, exist_ok=True)
        copied = copy_templates_to_folder(dest)
        messagebox.showinfo("Export Successful", f"Exported {len(copied)} templates to:\n{dest}")
        try:
            subprocess.Popen(["explorer.exe", str(dest)])
        except Exception:
            pass
        self._refresh_data_hub_list()

    def _refresh_data_hub_list(self) -> None:
        self.hub_listbox.delete("1.0", "end")
        lines = [
            "=== BUNDLED SYSTEM TEMPLATES ===",
            "  1. Declaration Template (PACS 6 & District metadata)",
            "  2. ShortTermLoans Template (17 rules, KCC & Crops)",
            "  3. Membership Template (PACS Village + 28 Rules)",
            "  4. LandDetails Template (Mandal & Village Khasra)",
            "  5. Savings Template (Savings Deposit Accounts)",
            "  6. LoanTransaction Template (Ledger Folio & Disbursals)",
            "",
            "=== CURRENT ACTIVE CONFIGURATION ===",
            f"  Master File:  {self.sheet_path.name if self.sheet_path else 'None (Default Sample)'}",
            f"  Target PACS:  {self.pacs_var.get()}",
            f"  District:     {self.district_var.get()}",
            f"  Portal URL:   {self.url_var.get()}",
            f"  Output Dir:   C:\\Users\\DEEPAK PAL\\Desktop\\DCT -1",
        ]
        self.hub_listbox.insert("end", "\n".join(lines) + "\n")

    # -------------------------------------------------------------------------
    # VIEW 5: ADMIN SETTINGS (Image 2 - Spectrum Pastel)
    # -------------------------------------------------------------------------
    def _build_admin_settings_view(self, parent: ctk.CTkFrame) -> None:
        self._build_page_title_bar(
            parent,
            title="Admin & automation settings",
            subtitle="Manage your default PACS and portal preferences."
        )

        content = ctk.CTkFrame(parent, fg_color="transparent")
        content.pack(fill="both", expand=True)
        content.grid_columnconfigure(0, weight=5)
        content.grid_columnconfigure(1, weight=4)

        # Left Card: Default configuration
        left_card = ctk.CTkFrame(
            content, fg_color=Palette.PANEL_BG, corner_radius=14,
            border_color=Palette.PANEL_BORDER, border_width=1
        )
        left_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        h_set = ctk.CTkFrame(left_card, fg_color="transparent")
        h_set.pack(fill="x", padx=20, pady=(18, 4))
        ctk.CTkLabel(h_set, text="⚙️", font=font_tnr(16)).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(h_set, text="Default configuration", font=font_tnr(15, weight="bold"), text_color=Palette.TEXT_TITLE).pack(side="left")

        ctk.CTkLabel(
            left_card, text="Set your default PACS, portal gateway and district for automation.",
            font=font_tnr(10), text_color=Palette.TEXT_MUTED
        ).pack(anchor="w", padx=20, pady=(0, 16))

        # Fields
        f_box = ctk.CTkFrame(left_card, fg_color="transparent")
        f_box.pack(fill="x", padx=20, pady=(0, 16))

        ctk.CTkLabel(f_box, text="Default PACS ID code", font=font_tnr(11, weight="bold"), text_color=Palette.TEXT_TITLE).pack(anchor="w", pady=(0, 4))
        ctk.CTkEntry(
            f_box, textvariable=self.pacs_var, height=36, corner_radius=8,
            font=font_tnr(11), fg_color=Palette.INPUT_BG,
            border_color=Palette.INPUT_BORDER, text_color=Palette.TEXT_TITLE
        ).pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(f_box, text="Default portal gateway", font=font_tnr(11, weight="bold"), text_color=Palette.TEXT_TITLE).pack(anchor="w", pady=(0, 4))
        ctk.CTkComboBox(
            f_box, values=PORTAL_URLS, variable=self.url_var, height=36, corner_radius=8,
            font=font_tnr(11), fg_color=Palette.INPUT_BG,
            border_color=Palette.INPUT_BORDER, button_color="#2563EB", text_color=Palette.TEXT_TITLE
        ).pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(f_box, text="Default district", font=font_tnr(11, weight="bold"), text_color=Palette.TEXT_TITLE).pack(anchor="w", pady=(0, 4))
        ctk.CTkComboBox(
            f_box, values=UP_DISTRICTS, variable=self.district_var, height=36, corner_radius=8,
            font=font_tnr(11), fg_color=Palette.INPUT_BG,
            border_color=Palette.INPUT_BORDER, button_color="#2563EB", text_color=Palette.TEXT_TITLE
        ).pack(fill="x", pady=(0, 20))

        def _on_save_prefs():
            save_persisted_config({
                "url": self.url_var.get().strip(),
                "login_id": self.login_var.get().strip(),
                "password": self.password_var.get(),
                "pacs_id": self.pacs_var.get().strip(),
                "district": self.district_var.get().strip(),
            })
            messagebox.showinfo("Saved", "Preferences saved successfully!")
            self._switch_view("setup")

        VisualAssets.create_gradient_button(
            left_card, text="  💾  Save preferences & return to setup  ",
            command=_on_save_prefs,
            width=320, height=44, corner_radius=10,
            font=font_main(12, weight="bold")
        ).pack(fill="x", padx=20, pady=(0, 20))

        # Right Card: Brand Showcase Card matching Image 2 & Section 17
        right_card = ctk.CTkFrame(
            content, fg_color=("#F8FAFF", "#11223B"), corner_radius=14,
            border_color=Palette.PANEL_BORDER, border_width=1
        )
        right_card.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        brand_inner = ctk.CTkFrame(right_card, fg_color="transparent")
        brand_inner.place(relx=0.5, rely=0.5, anchor="center")

        # Big "D1" Icon Badge
        ctk.CTkButton(
            brand_inner, text="D1", width=72, height=72, corner_radius=18,
            fg_color="#2454E6", hover_color="#1D45C8", text_color="#FFFFFF",
            font=font_tnr(28, weight="bold")
        ).pack(pady=(0, 16))

        ctk.CTkLabel(
            brand_inner, text="DCTtool v2.0 Enterprise",
            font=font_tnr(20, weight="bold"),
            text_color=Palette.TEXT_HEADING
        ).pack()

        ctk.CTkLabel(
            brand_inner, text="PACS Master Automation Platform",
            font=font_tnr(13),
            text_color=Palette.TEXT_SECONDARY
        ).pack(pady=(4, 16))

        ctk.CTkFrame(brand_inner, fg_color=Palette.BORDER_DIVIDER, height=1, width=220).pack(pady=(0, 16))

        ctk.CTkLabel(
            brand_inner, text="Developed by Deepak Pal",
            font=font_tnr(13, weight="bold"),
            text_color=Palette.TEXT_PRIMARY
        ).pack()

        ctk.CTkLabel(
            brand_inner, text="© 2026 Deepak Pal. All Rights Reserved.",
            font=font_tnr(11),
            text_color=Palette.TEXT_MUTED
        ).pack(pady=(4, 0))

    # -------------------------------------------------------------------------
    # Core Actions & Master Engine Execution (Preserved 100%)
    # -------------------------------------------------------------------------
    def _browse_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Master DCT Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        if path:
            self.sheet_path = Path(path)
            self.file_title_var.set(self.sheet_path.name)
            self.file_sub_var.set(f"File loaded successfully\n{self.sheet_path.parent}")
            self.drop_card.configure(border_color="#10B981")
            if hasattr(self, "btn_remove_file"):
                self.btn_remove_file.pack(side="left", fill="x", expand=True, padx=(2, 0))
            self._log(f"[INFO] Selected Master File: {self.sheet_path.name}")

            try:
                import openpyxl
                meta_wb = openpyxl.load_workbook(self.sheet_path, data_only=True, read_only=True)
                clean_pacs, dist, pid = extract_pacs_metadata_from_workbook(meta_wb, fallback_stem=self.sheet_path.stem)
                meta_wb.close()
                if clean_pacs:
                    self.target_pacs_label_var.set(f"[Cite: 1] ({clean_pacs})")
                    self._log(f"[METADATA] Auto-detected PACS Name: {clean_pacs}")
                if dist:
                    self.district_var.set(dist)
                    self._log(f"[METADATA] Auto-detected District: {dist}")
                if pid:
                    self.pacs_var.set(pid)
                    self._log(f"[METADATA] Auto-detected PACS ID: {pid}")
            except Exception as meta_err:
                self._log(f"[WARN] Metadata auto-detection note: {meta_err}")

            self._refresh_data_hub_list()

    def _remove_file(self) -> None:
        self.sheet_path = None
        self.file_title_var.set("No master sheet loaded")
        self.file_sub_var.set("Click or drag & drop to load master Excel file (.xlsx)")
        self.drop_card.configure(border_color=Palette.DROPZONE_BORDER)
        if hasattr(self, "btn_remove_file"):
            self.btn_remove_file.pack_forget()
        self._log("[FILE] Master sheet removed by user.")

    def _download_sample_templates(self) -> None:
        dest = filedialog.askdirectory(title="Select Folder to Save Templates")
        if dest:
            dest_p = Path(dest)
            copied = copy_templates_to_folder(dest_p)
            messagebox.showinfo("Templates Copied", f"Copied {len(copied)} templates to:\n{dest_p}")
            self._log(f"[INFO] Copied {len(copied)} templates to {dest_p}")
            try:
                subprocess.Popen(["explorer.exe", str(dest_p)])
            except Exception:
                pass

    def _save_logs(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")],
            initialfile=f"dcttool_log_{datetime.date.today().isoformat()}.txt"
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.log_textbox.get("1.0", "end"))
            messagebox.showinfo("Logs Saved", f"Log file saved successfully:\n{path}")
            self._log(f"[LOGS] Exported logs to: {path}")

    def _clear_logs(self) -> None:
        self._clear_consoles()
        self._log("[LOG] Console logs cleared by user.")

    def _log(self, text: str) -> None:
        self._log_queue.put(text)

    def _clear_consoles(self) -> None:
        if hasattr(self, "full_log_textbox"):
            self.full_log_textbox.delete("1.0", "end")
            self.full_log_textbox.insert("end", "\n\n\n\n               [ >_ ]\n            No logs yet\n   Start the master process to view execution logs.\n")
        if hasattr(self, "log_textbox"):
            self.log_textbox.delete("1.0", "end")

    def _drain_log(self) -> None:
        while not self._log_queue.empty():
            msg = self._log_queue.get_nowait()
            if hasattr(self, "log_textbox"):
                self.log_textbox.insert("end", msg + "\n")
                self.log_textbox.see("end")
            if hasattr(self, "full_log_textbox"):
                current_full = self.full_log_textbox.get("1.0", "end")
                if "No logs yet" in current_full:
                    self.full_log_textbox.delete("1.0", "end")
                self.full_log_textbox.insert("end", msg + "\n")
                self.full_log_textbox.see("end")
        self.after(100, self._drain_log)

    def _start_process_thread(self) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showwarning("In Progress", "A processing operation is already running.")
            return

        if not self.sheet_path or not self.sheet_path.exists():
            messagebox.showerror("No Master File", "Please select a valid Master Excel Sheet (.xlsx) first.\n\nClick 'Upload' or 'Replace file' to select your file.")
            return

        if self.chk_membership.get():
            login = self.login_var.get().strip()
            if not login or login == "demo@example.com":
                messagebox.showerror(
                    "Portal Login Required",
                    "Membership template create karne ke liye sahi Portal Login ID aur Password zaroori hai!\n\n"
                    "Bina PACS portal login ke genuine dropdown gaon prapt nahi ho sakte aur upload karne par sabhi members 'Others' ho jate hain.\n\n"
                    "Kripya Card 2 (Portal Configuration) me apna PACS Login ID aur Password check karein."
                )
                return

        self.btn_start.configure(state="disabled", text="PROCESSING...")
        self.progress_pct_var.set("10%")
        self.progress_role_var.set("● Running")
        self.progress_bar.set(0.1)

        def worker():
            try:
                self._execute_process()
            except Exception as e:
                self._log(f"[FATAL ERROR] {e}")
                self.after(0, lambda: messagebox.showerror("Processing Failed", str(e)))
            finally:
                self.after(0, lambda: self.btn_start.configure(state="normal", text="  ▶  Start process  "))
                self.after(0, lambda: self.progress_role_var.set("● Ready"))
                self.after(0, lambda: self.progress_pct_var.set("100%"))
                self.after(0, lambda: self.progress_bar.set(1.0))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def _execute_process(self) -> None:
        try:
            self._log("\n" + "="*56)
            self._log(f"[START] Starting Master Process for: {self.sheet_path.name}")
            self._log(f"[CONFIG] PACS ID: {self.pacs_var.get()} | District: {self.district_var.get()}")
            self._log(f"[CONFIG] Portal URL: {self.url_var.get()}")

            save_persisted_config({
                "url": self.url_var.get().strip(),
                "login_id": self.login_var.get().strip(),
                "password": self.password_var.get(),
                "pacs_id": self.pacs_var.get().strip(),
                "district": self.district_var.get().strip(),
            })

            settings = MasterSettings(
                seed_path=self.sheet_path,
                url=self.url_var.get().strip(),
                login_id=self.login_var.get().strip(),
                password=self.password_var.get(),
                pacs_id=self.pacs_var.get().strip(),
                district=self.district_var.get().strip(),
                templates=TemplateSelection(
                    declaration=self.chk_declaration.get(),
                    membership=self.chk_membership.get(),
                    stloans=self.chk_stloans.get(),
                    land=self.chk_land.get(),
                    loan_transaction=self.chk_loan.get(),
                    savings=self.chk_savings.get(),
                    dropdown_list=self.chk_dropdown.get(),
                    stloans_to_lt=self.chk_stloans_to_lt.get(),
                ),
            )

            def on_status(msg: str):
                self._log(f"[ENGINE] {msg}")
                low = msg.lower()
                if "loading master" in low:
                    self.after(0, lambda: [self.progress_pct_var.set("20%"), self.progress_bar.set(0.2)])
                elif "portal" in low:
                    self.after(0, lambda: [self.progress_pct_var.set("35%"), self.progress_bar.set(0.35)])
                elif "short term loans" in low or "shorttermloans" in low:
                    self.after(0, lambda: [self.progress_pct_var.set("50%"), self.progress_bar.set(0.5)])
                elif "loan transaction" in low or "loantransaction" in low:
                    self.after(0, lambda: [self.progress_pct_var.set("65%"), self.progress_bar.set(0.65)])
                elif "membership" in low:
                    self.after(0, lambda: [self.progress_pct_var.set("80%"), self.progress_bar.set(0.8)])
                elif "saving" in low or "export" in low:
                    self.after(0, lambda: [self.progress_pct_var.set("90%"), self.progress_bar.set(0.9)])

            result = process_master_workbook(settings, on_status=on_status)
            self._processed_wb = result.workbook
            self._process_report = result.report

            # Destination folder: Desktop\DCT -1\<CLEAN_PACS_NAME> (ALL CAPS from Declaration sheet)
            clean_pacs, _, _ = extract_pacs_metadata_from_workbook(result.workbook, fallback_stem=settings.seed_path.stem)
            folder_name = clean_pacs if clean_pacs and clean_pacs != "PACS" else settings.seed_path.stem
            desktop_base = Path.home() / "Desktop" / "DCT -1"
            out_dir = desktop_base / folder_name
            out_dir.mkdir(parents=True, exist_ok=True)
            self._log(f"[TARGET] Export destination folder: {out_dir}")

            exported = save_split_templates(
                result.workbook,
                out_dir,
                settings.templates,
                report=result.report,
                original_path=settings.seed_path,
                on_status=on_status,
            )

            self.after(0, lambda: [self.progress_pct_var.set("100%"), self.progress_bar.set(1.0)])
            self._log(f"\n[SUCCESS] Successfully exported {len(exported)} template(s) into:")
            self._log(f"  {out_dir}")
            for p in exported:
                self._log(f"  [OK] {p.name}")

            # Completion Sound Beep
            try:
                import winsound
                winsound.Beep(800, 200)
                winsound.Beep(1000, 200)
                winsound.Beep(1200, 400)
            except Exception:
                pass

            # Open folder in Windows Explorer
            try:
                subprocess.Popen(["explorer.exe", str(out_dir)])
            except Exception:
                pass

            file_list_str = "\n".join(f"  ✔  {p.name}" for p in exported)
            success_msg = (
                f"🎉 Process Successfully Completed!\n\n"
                f"🏛️ PACS Name: {clean_pacs}\n"
                f"📁 Output Folder: {out_dir}\n\n"
                f"Generated Ready Templates:\n{file_list_str}\n\n"
                f"All business rules and portal requirements were successfully applied."
            )
            self.after(0, lambda msg=success_msg: messagebox.showinfo("Successfully Completed", msg))
            self._refresh_data_hub_list()

        except Exception as e:
            err_msg = str(e).strip() if str(e).strip() else repr(e)
            self._log(f"\n[ERROR] Processing Failed: {err_msg}")
            import traceback
            trace_str = traceback.format_exc()
            self._log(f"[TRACEBACK]\n{trace_str}")
            error_popup_msg = (
                f"⚠️ Process Failed / Error Alert\n\n"
                f"Karan / Error Details:\n{err_msg}\n\n"
                f"Kripya ye check karein:\n"
                f"1. Agar koi Excel template open hai to use band karein.\n"
                f"2. Internet connection aur Portal URL verify karein.\n"
                f"3. Live Process Log me step-by-step detail check karein."
            )
            self.after(0, lambda msg=error_popup_msg: messagebox.showerror("Process Error", msg))


def main() -> None:
    app = DCTtoolApp()
    app.mainloop()


if __name__ == "__main__":
    main()
