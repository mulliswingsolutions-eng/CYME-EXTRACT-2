# src/app_tk.py  — Premium UI + Island workflow (click row to activate)
from __future__ import annotations
import os, sys, json, threading, queue, traceback, platform, subprocess
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, TYPE_CHECKING, cast

import customtkinter as ctk
import pandas as pd

# ----- Project setup ----------------------------------------------------------
# ----- Project setup ----------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

def resource_path(*parts: str) -> Path:
    """
    Return a Path to packaged resources that works in dev and PyInstaller.
    Example: resource_path('icons', 'nature.ico')
    """
    if getattr(sys, "frozen", False):
        # Pylance-safe access to _MEIPASS without tripping attribute checks
        base = Path(cast(str, getattr(sys, "_MEIPASS", str(BASE_DIR.parent))))
        return base.joinpath(*parts)
    return BASE_DIR.parent.joinpath(*parts)

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

# Pipeline pieces
from Modules.IslandChecker import analyze_and_set_island_context
from Modules.General import write_general_sheet, get_island_context
from Modules.Pins import write_pins_sheet
from Modules.Bus import write_bus_sheet
from Modules.Voltage_Source import write_voltage_source_sheet
from Modules.Load import write_load_sheet
from Modules.Line import write_line_sheet
from Modules.Transformer import write_transformer_sheet
from Modules.Switch import write_switch_sheet
from Modules.Shunt import write_shunt_sheet

APP_NAME = "CYME → XLSX Extractor"
CONF_PATH = Path.home() / ".cyme_extractor_gui.json"

DEFAULT_SHEETS = {
    "General": True,
    "Pins": True,
    "Bus": True,
    "Voltage Source": True,
    "Load": True,
    "Line": True,
    "Transformer": True,
    "Switch": True,
    "Shunt": True,
}

# ----- Utilities --------------------------------------------------------------

def _filter_context_to_island(self, active_island: int):
    """Keep only the selected island in context (used right before writing)."""
    ctx = get_island_context() or {}
    bus_to_island: dict[str, int] = dict(ctx.get("bus_to_island", {}))
    slack_per_island: dict[int, str] = dict(ctx.get("slack_per_island", {}))
    bad_buses_existing: set[str] = set(ctx.get("bad_buses", set()))

    # everything not in the chosen island gets marked bad
    bad_buses_new = {b for b, isl in bus_to_island.items() if isl != active_island}

    new_ctx = {
        "bus_to_island": {b: i for b, i in bus_to_island.items() if i == active_island},
        "slack_per_island": {active_island: slack_per_island.get(active_island, "")},
        "bad_buses": set(bad_buses_existing) | set(bad_buses_new),
    }

    import Modules.General as G  # type: ignore
    if hasattr(G, "set_island_context"):
        G.set_island_context(new_ctx)  # type: ignore[attr-defined]
    elif hasattr(G, "_ISLAND_CONTEXT"):
        setattr(G, "_ISLAND_CONTEXT", new_ctx)  # type: ignore[attr-defined]
    elif hasattr(G, "ISLAND_CONTEXT"):
        setattr(G, "ISLAND_CONTEXT", new_ctx)  # type: ignore[attr-defined]
    else:
        def _fake_get_island_context():
            return new_ctx
        G.get_island_context = _fake_get_island_context  # type: ignore[assignment]


def _keep_sourceful_islands_context(self, in_path: Path):
    """
    No active selection:
    - Keep ALL islands in context for UI.
    - Mark ONLY the islands that do not contain a Substation/VS-page source as bad.
    """
    from Modules.General import get_island_context, set_island_context, safe_name
    import xml.etree.ElementTree as ET

    ctx = get_island_context() or {}
    islands: dict[int, set[str]] = dict(ctx.get("islands", {}))
    bus_to_island: dict[str, int] = dict(ctx.get("bus_to_island", {}))
    slack_per_island: dict[int, str] = dict(ctx.get("slack_per_island", {}))

    # If the IslandChecker didn't include "islands", synthesize from bus_to_island.
    if not islands and bus_to_island:
        for b, isl in bus_to_island.items():
            islands.setdefault(isl, set()).add(b)

    # --- Find real Voltage Source page buses from the input file ---
    try:
        root = ET.fromstring(Path(in_path).read_text(encoding="utf-8", errors="ignore"))
    except Exception:
        root = ET.Element("Empty")

    vs_nodes: set[str] = set()
    for topo in root.findall(".//Topo"):
        ntype = (topo.findtext("NetworkType") or "").strip().lower()
        eq_mode = (topo.findtext("EquivalentMode") or "").strip()
        if ntype != "substation" or eq_mode == "1":
            continue
        srcs = topo.find("./Sources")
        if srcs is None:
            continue
        for src in srcs.findall("./Source"):
            nid = safe_name(src.findtext("SourceNodeID") or "")
            if nid:
                vs_nodes.add(nid)

    # Islands that intersect VS-page nodes are sourceful (good)
    sourceful_islands = {i for i, buses in islands.items() if any(b in vs_nodes for b in buses)}
    bad_islands = set(islands.keys()) - sourceful_islands

    # Keep full mapping for the UI
    full_bus_to_island = {b: i for i, buses in islands.items() for b in buses}
    bad_buses = set().union(*(islands[i] for i in bad_islands)) if islands else set()

    # Preserve slacks if they exist; we don't rely on them for "good" anymore
    new_ctx = {
        "bus_to_island": full_bus_to_island,
        "slack_per_island": slack_per_island,
        "bad_buses": bad_buses,
        "islands": islands,
    }
    set_island_context(new_ctx)



def load_conf() -> dict:
    try:
        return json.loads(CONF_PATH.read_text())
    except Exception:
        return {}

def save_conf(data: dict) -> None:
    """Merge and persist small UI settings (best-effort, safe)."""
    try:
        current = {}
        try:
            current = json.loads(CONF_PATH.read_text())
        except Exception:
            current = {}
        current.update(data)
        CONF_PATH.write_text(json.dumps(current, indent=2))
    except Exception:
        pass

def open_in_file_explorer(path: Path) -> None:
    p = str(path)
    if platform.system() == "Windows":
        os.startfile(p)  # type: ignore[attr-defined]
    elif platform.system() == "Darwin":
        subprocess.call(["open", p])
    else:
        subprocess.call(["xdg-open", p])

def set_window_icon(window: tk.Tk | ctk.CTk) -> tk.PhotoImage | None:
    """
    Window/taskbar icon:
      - Windows: prefer nature.ico (iconbitmap)
      - Else: use nature.png if present (iconphoto)
    Works with PyInstaller via resource_path().
    """
    try:
        # Prefer new branded assets
        ico = resource_path("icons", "cyme_logo.ico")
        png = resource_path("icons", "cyme_logo_light.png")
        old_ico = resource_path("icons", "nature.ico")
        old_png = resource_path("icons", "nature.png")

        if platform.system() == "Windows" and ico.exists():
            window.iconbitmap(default=str(ico))
            return None

        if png.exists():
            img = tk.PhotoImage(file=str(png))
            window.iconphoto(True, img)
            return img  # caller must keep a reference

        # Fallback to legacy files if new ones are missing
        if platform.system() == "Windows" and old_ico.exists():
            window.iconbitmap(default=str(old_ico))
            return None
        if old_png.exists():
            img = tk.PhotoImage(file=str(old_png))
            window.iconphoto(True, img)
            return img
    except Exception:
        pass
    return None

def _hex(c: str) -> tuple[int, int, int]:
    c = c.lstrip('#')
    return int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)

def ensure_brand_assets() -> None:
    """Generate light/dark PNG logos and ICO if missing.
    Creates:
      - icons/cyme_logo_light.png
      - icons/cyme_logo_dark.png
      - icons/cyme_logo.ico
    The design: pastel ring, inner disc, clean bolt, drawn at high res and downsampled.
    """
    try:
        from PIL import Image, ImageDraw, ImageFilter  # type: ignore
    except Exception:
        return

    icons_dir = resource_path("icons")
    icons_dir.mkdir(parents=True, exist_ok=True)

    out_light = icons_dir / "cyme_logo_light.png"
    out_dark = icons_dir / "cyme_logo_dark.png"
    out_ico = icons_dir / "cyme_logo.ico"

    def make_variant(mode: str, path: Path) -> Image.Image:
        W = 1024
        img = Image.new("RGBA", (W, W), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        if mode == "dark":
            bg = _hex("#0F172A")
            ring = _hex("#C3B5FD")
            bolt = _hex("#FFFFFF")
            glow = (167, 139, 250, 90)
        else:
            bg = _hex("#FFFFFF")
            ring = _hex("#A78BFA")
            bolt = _hex("#6D28D9")
            glow = (167, 139, 250, 70)

        cx = cy = W // 2
        R = int(W * 0.42)
        ring_w = int(W * 0.08)

        # Drop shadow
        shadow = Image.new("RGBA", (W, W), (0, 0, 0, 0))
        sd = ImageDraw.Draw(shadow)
        sd.ellipse((cx - R, cy - R, cx + R, cy + R), fill=(0, 0, 0, 160))
        shadow = shadow.filter(ImageFilter.GaussianBlur(radius=int(W * 0.03)))
        img.alpha_composite(shadow)

        # Outer ring and subtle highlight
        draw.ellipse((cx - R, cy - R, cx + R, cy + R), outline=ring + (255,), width=ring_w)
        inset = int(ring_w * 0.35)
        draw.ellipse((cx - R + inset, cy - R + inset, cx + R - inset, cy + R - inset),
                     outline=(255, 255, 255, 60), width=max(1, int(ring_w * 0.25)))

        # Inner disc
        r2 = R - int(ring_w * 0.75)
        draw.ellipse((cx - r2, cy - r2, cx + r2, cy + r2), fill=bg + (255,))

        # Soft radial glow
        glow_im = Image.new("RGBA", (W, W), (0, 0, 0, 0))
        gd = ImageDraw.Draw(glow_im)
        gd.ellipse((cx - int(r2 * 0.95), cy - int(r2 * 0.95), cx + int(r2 * 0.95), cy + int(r2 * 0.95)), fill=glow)
        glow_im = glow_im.filter(ImageFilter.GaussianBlur(radius=int(W * 0.05)))
        img.alpha_composite(glow_im)

        # --- File conversion + power systems hint ---
        # Document sheet
        doc_w, doc_h = int(r2 * 1.05), int(r2 * 1.05)
        dx0, dy0 = cx - int(doc_w * 0.55), cy - int(doc_h * 0.50)
        dx1, dy1 = cx + int(doc_w * 0.15), cy + int(doc_h * 0.45)
        doc_fill = (255, 255, 255, 235) if mode == "light" else (20, 22, 26, 235)
        doc_border = ring + (200,)
        try:
            draw.rounded_rectangle((dx0, dy0, dx1, dy1), radius=int(r2 * 0.12), fill=doc_fill, outline=doc_border, width=max(2, int(W * 0.004)))
        except Exception:
            draw.rectangle((dx0, dy0, dx1, dy1), fill=doc_fill, outline=doc_border, width=max(2, int(W * 0.004)))

        # Folded corner
        fold = int(min(dx1-dx0, dy1-dy0) * 0.18)
        fold_poly = [(dx1-fold, dy0), (dx1, dy0), (dx1, dy0+fold)]
        draw.polygon(fold_poly, fill=(255,255,255,200) if mode=="light" else (34,36,42,220))
        draw.line([(dx1-fold, dy0), (dx1, dy0), (dx1, dy0+fold)], fill=doc_border, width=max(1, int(W*0.002)))

        # Spreadsheet grid hint inside document
        gcol = (ring[0], ring[1], ring[2], 120)
        for i in range(1,4):
            x = dx0 + int((dx1-dx0) * (i/4.5))
            draw.line([(x, dy0+int((dy1-dy0)*0.20)), (x, dy1-int((dy1-dy0)*0.15))], fill=gcol, width=max(1, int(W*0.002)))
        for j in range(1,5):
            y = dy0 + int((dy1-dy0) * (j/6.0))
            draw.line([(dx0+int((dx1-dx0)*0.07), y), (dx1-int((dx1-dx0)*0.07), y)], fill=gcol, width=max(1, int(W*0.002)))

        # Conversion arrow (to the right of the sheet)
        ax0 = dx1 + int(r2*0.05); ay0 = cy - int(r2*0.10)
        ax1 = dx1 + int(r2*0.40); ay1 = cy + int(r2*0.10)
        draw.rounded_rectangle((ax0, ay0, ax1, ay1), radius=int(r2*0.05), fill=(ring[0],ring[1],ring[2],180))
        tri = [(ax1-int(r2*0.02), cy), (ax1-int(r2*0.10), ay0), (ax1-int(r2*0.10), ay1)]
        draw.polygon(tri, fill=(ring[0],ring[1],ring[2],200))

        # Power system hint: small 3-node network at bottom-left of sheet
        nx = dx0 + int((dx1-dx0)*0.22)
        ny = dy1 - int((dy1-dy0)*0.20)
        r = max(2, int(W*0.01))
        nodes = [(nx, ny), (nx-int(r2*0.12), ny-int(r2*0.08)), (nx+int(r2*0.12), ny-int(r2*0.10))]
        # lines
        draw.line([nodes[0], nodes[1]], fill=doc_border, width=max(2, int(W*0.004)))
        draw.line([nodes[0], nodes[2]], fill=doc_border, width=max(2, int(W*0.004)))
        for (px,py) in nodes:
            draw.ellipse((px-r,py-r,px+r,py+r), fill=doc_border)

        out = img.resize((256, 256), Image.LANCZOS)
        out.save(path)
        return out

    try:
        if not out_light.exists():
            make_variant("light", out_light)
        if not out_dark.exists():
            make_variant("dark", out_dark)
        if not out_ico.exists():
            # Build multi-size ICO from the light variant
            im = Image.open(out_light) if out_light.exists() else make_variant("light", out_light)
            im.save(out_ico, sizes=[(256,256),(128,128),(64,64),(32,32),(16,16)])
    except Exception:
        pass

def setup_appearance(initial_mode: str = "light") -> tuple[str, str, int, int, dict, str]:
    """CustomTkinter global theme + fonts + colors (single light theme)."""
    # Force light mode
    mode = "light"
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    # Cross-platform typefaces
    if platform.system() == "Darwin":
        ui_font = "SF Pro Text"
        mono_font = "Menlo"
    elif platform.system() == "Windows":
        ui_font = "Segoe UI Variable"
        mono_font = "Cascadia Mono"
    else:
        ui_font = "Inter"
        mono_font = "Liberation Mono"

    ui_size = 12
    header_size = 18

    def palette() -> dict:
        # Light theme only
        return {
            "BG": "#F7F8FA",
            "CARD": "#FFFFFF",
            "TEXT": "#0B0F19",
            "MUTED": "#6B7280",
            "BORDER": "#E5E7EB",
            "ACCENT": "#A78BFA",
            "ACCENT_HOVER": "#8B5CF6",
            "ACCENT_SOFT": "#F1EDFE",
            "ACCENT_SOFT_HOVER": "#E6E0FD",
            "DANGER": "#F43F5E",
            "DANGER_HOVER": "#E11D48",
            "CONSOLE_BG": "#0F172A",
            "CONSOLE_FG": "#E5E7EB",
            "INPUT_BG": "#FFFFFF",
            "INPUT_FG": "#0B0F19",
            "INPUT_BORDER": "#D1D5DB",
        }

    colors = palette()
    return ui_font, mono_font, ui_size, header_size, colors, mode

# ----- App --------------------------------------------------------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        try:
            ensure_brand_assets()
        except Exception:
            pass
        self._icon_ref = set_window_icon(self)  # keep ref
        self.title(APP_NAME)
        self.geometry("1180x720")

        # theme (light only)
        conf = load_conf()
        self.UI_FONT, self.MONO_FONT, self.UI_SIZE, self.HEADER_SIZE, self.COL, _ = setup_appearance("light")
        self.configure(fg_color=self.COL["BG"])

        # state
        self.events: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self.in_path = tk.StringVar(value=conf.get("last_input", ""))
        self.out_path = tk.StringVar(value=conf.get("last_output", str(Path.cwd() / "CYME_Extract.xlsx")))
        self.sheet_vars: dict[str, tk.BooleanVar] = {name: tk.BooleanVar(value=DEFAULT_SHEETS[name]) for name in DEFAULT_SHEETS}
        self.active_island_id: int | None = None  # click-to-activate
        self._suppress_island_event: bool = False  # guard to prevent feedback loops
        self._handling_island_click: bool = False

        # UI
        self._build_header()
        self._build_tabs()
        # (No dark mode — single palette already applied)
        self._poll_events()

    # -- header
    def _build_header(self):
        self.header = ctk.CTkFrame(self, fg_color=self.COL["BG"], corner_radius=0)
        # tighter top spacing
        self.header.pack(fill="x", padx=24, pady=(0, 0))
        try:
            self.header.pack_propagate(False)
            self.header.configure(height=40)
        except Exception:
            pass

        # --- brand: logo (light only) ---
        from PIL import Image  # type: ignore
        logo_png = resource_path("icons", "cyme_logo_light.png")
        if not logo_png.exists():
            logo_png = resource_path("icons", "nature.png")
        try:
            self.brand_img = ctk.CTkImage(light_image=Image.open(str(logo_png)), size=(28, 28))
            self.brand_img_label = ctk.CTkLabel(self.header, image=self.brand_img, text="", fg_color=self.COL["BG"])
            self.brand_img_label.pack(side="left", padx=(2, 8), pady=0)
        except Exception:
            pass


        title = ctk.CTkLabel(self.header, text=APP_NAME, font=(self.UI_FONT, self.HEADER_SIZE, "bold"), text_color=self.COL["TEXT"]) 
        title.pack(side="left", pady=0)

        # Right side spacer (no theme switch)
        right = ctk.CTkFrame(self.header, fg_color=self.COL["BG"])
        right.pack(side="right")

    def _build_theme_switch(self, parent: ctk.CTkFrame):
        # Build small sun/moon icons (Pillow), with fallback to text if Pillow missing
        self.sun_icon = None
        self.moon_icon = None
        try:
            from PIL import Image, ImageDraw, ImageTk  # type: ignore
            def sun_img(color: str) -> tk.PhotoImage:
                W = 18
                im = Image.new("RGBA", (W, W), (0,0,0,0))
                d = ImageDraw.Draw(im)
                # core
                d.ellipse((4,4,14,14), fill=color)
                # rays
                for a in range(0,360,45):
                    # tiny rectangles as rays
                    d.pieslice((1,1,17,17), a-2, a+2, fill=color)
                return ImageTk.PhotoImage(im)
            def moon_img(color: str) -> tk.PhotoImage:
                W = 18
                im = Image.new("RGBA", (W, W), (0,0,0,0))
                d = ImageDraw.Draw(im)
                d.ellipse((3,3,15,15), fill=color)
                d.ellipse((7,3,17,15), fill=(0,0,0,0))
                return ImageTk.PhotoImage(im)
            self.sun_icon = sun_img(self.COL["TEXT"])
            self.moon_icon = moon_img(self.COL["TEXT"])
        except Exception:
            pass

        row = ctk.CTkFrame(parent, fg_color=self.COL["BG"])
        row.pack(side="right")
        self.sun_lbl = ctk.CTkLabel(row, text="", image=self.sun_icon, text_color=self.COL["TEXT"]) if self.sun_icon else ctk.CTkLabel(row, text="☀", text_color=self.COL["TEXT"]) 
        self.sun_lbl.pack(side="left", padx=(0,6))
        self.theme_bool = tk.IntVar(value=1 if self.theme_mode == "dark" else 0)
        self.theme_switch = ctk.CTkSwitch(row, text="", command=self._on_theme_switch,
                                          variable=self.theme_bool, onvalue=1, offvalue=0,
                                          width=52, height=28, progress_color=self.COL["ACCENT"])
        self.theme_switch.pack(side="left")
        self.moon_lbl = ctk.CTkLabel(row, text="", image=self.moon_icon, text_color=self.COL["TEXT"]) if self.moon_icon else ctk.CTkLabel(row, text="🌙", text_color=self.COL["TEXT"]) 
        self.moon_lbl.pack(side="left", padx=(6,0))

    # -- tabs
    def _build_tabs(self):
        self.tabs = ctk.CTkTabview(self, fg_color=self.COL["BG"],
                                   segmented_button_selected_color=self.COL["ACCENT"],
                                   segmented_button_selected_hover_color=self.COL["ACCENT_HOVER"])
        # reduce space above segmented tabs further
        self.tabs.pack(fill="both", expand=True, padx=24, pady=(0, 2))
        try:
            seg = self.tabs._segmented_button
            seg.configure(font=(self.UI_FONT, self.UI_SIZE + 2), height=34, corner_radius=10)
            # try to remove internal padding regardless of geometry manager
            try:
                seg.grid_configure(pady=0, ipady=0)
            except Exception:
                pass
            try:
                seg.pack_configure(pady=0, ipady=0)
            except Exception:
                pass
        except Exception:
            pass

        self.tab_run = self.tabs.add("Run")
        self.tab_islands = self.tabs.add("Islands")

        self._build_run_tab(self.tab_run)
        self._build_islands_tab(self.tab_islands)

    # -- run tab
    def _build_run_tab(self, parent: ctk.CTkFrame):
        # card: file inputs
        self.run_card_file = ctk.CTkFrame(parent, fg_color=self.COL["CARD"], corner_radius=16)
        self.run_card_file.pack(fill="x", padx=4, pady=(6, 10))

        # input row
        row1 = ctk.CTkFrame(self.run_card_file, fg_color=self.COL["CARD"], corner_radius=0)
        row1.pack(fill="x", padx=18, pady=(18, 8))
        self.run_row1 = row1
        ctk.CTkLabel(row1, text="CYME File    ", font=(self.UI_FONT, self.UI_SIZE), text_color=self.COL["TEXT"]).pack(side="left", padx=(2, 12))
        self.in_entry = ctk.CTkEntry(row1, textvariable=self.in_path, height=38, font=(self.UI_FONT, self.UI_SIZE))
        # expand entry to fill available horizontal space
        self.in_entry.pack(side="left", padx=(0, 12), fill="x", expand=True)
        ctk.CTkButton(row1, text="Browse…", command=self._browse_in,
                      fg_color=self.COL["ACCENT"], hover_color=self.COL["ACCENT_HOVER"],
                      font=(self.UI_FONT, self.UI_SIZE), height=38, corner_radius=10).pack(side="left", padx=(0, 8))
        ctk.CTkButton(row1, text="Open Input Folder", command=self._open_in_dir,
                      fg_color="#EEF2FF", hover_color="#E0E7FF", text_color=self.COL["ACCENT"],
                      font=(self.UI_FONT, self.UI_SIZE), height=38, corner_radius=10).pack(side="left", padx=(0, 4))

        # output row
        row2 = ctk.CTkFrame(self.run_card_file, fg_color=self.COL["CARD"], corner_radius=0)
        row2.pack(fill="x", padx=18, pady=(6, 18))
        self.run_row2 = row2
        ctk.CTkLabel(row2, text="Output Excel", font=(self.UI_FONT, self.UI_SIZE), text_color=self.COL["TEXT"]).pack(side="left", padx=(2, 12))
        self.out_entry = ctk.CTkEntry(row2, textvariable=self.out_path, height=38, font=(self.UI_FONT, self.UI_SIZE))
        self.out_entry.pack(side="left", padx=(0, 12), fill="x", expand=True)
        ctk.CTkButton(row2, text="Save As…", command=self._browse_out,
                      fg_color=self.COL["ACCENT"], hover_color=self.COL["ACCENT_HOVER"],
                      font=(self.UI_FONT, self.UI_SIZE), height=38, corner_radius=10).pack(side="left", padx=(0, 8))
        ctk.CTkButton(row2, text="Open Output Folder", command=self._open_out_dir,
                      fg_color="#EEF2FF", hover_color="#E0E7FF", text_color=self.COL["ACCENT"],
                      font=(self.UI_FONT, self.UI_SIZE), height=38, corner_radius=10).pack(side="left", padx=(0, 4))

        # card: sheets with responsive layout
        self.checks_card = ctk.CTkFrame(parent, fg_color=self.COL["CARD"], corner_radius=16)
        self.checks_card.pack(fill="x", padx=4, pady=(6, 10))
        ctk.CTkLabel(self.checks_card, text="Sheets", font=(self.UI_FONT, self.UI_SIZE, "bold"), text_color=self.COL["TEXT"]).pack(anchor="w", padx=18, pady=(14, 0))

        self.checks_grid = ctk.CTkFrame(self.checks_card, fg_color=self.COL["CARD"])
        self.checks_grid.pack(fill="x", padx=12, pady=8)

        # create checkboxes once
        self._sheet_checks: list[ctk.CTkCheckBox] = []
        # slightly smaller font just for these checks
        check_font_size = max(10, self.UI_SIZE - 2)
        for name in DEFAULT_SHEETS:
            cb = ctk.CTkCheckBox(
                self.checks_grid,
                text=name,
                variable=self.sheet_vars[name],
                fg_color=self.COL["ACCENT"],
                hover_color=self.COL["ACCENT_HOVER"],
                font=(self.UI_FONT, check_font_size),
                text_color=self.COL["TEXT"],
            )
            self._sheet_checks.append(cb)
        # responsive layout
        self.checks_grid.bind("<Configure>", self._relayout_sheet_checks)
        self._relayout_sheet_checks()  # initial

        # actions
        self.run_actions = ctk.CTkFrame(parent, fg_color=self.COL["CARD"], corner_radius=16)
        self.run_actions.pack(fill="x", padx=4, pady=(6, 10))
        self.run_btn = ctk.CTkButton(self.run_actions, text="Process", command=self._start_run,
                                     fg_color=self.COL["ACCENT"], hover_color=self.COL["ACCENT_HOVER"],
                                     font=(self.UI_FONT, self.UI_SIZE), height=42, corner_radius=14)
        self.run_btn.pack(side="left", padx=18, pady=14)
        self.quit_btn = ctk.CTkButton(self.run_actions, text="Quit", command=self.destroy,
                                      fg_color=self.COL["DANGER"], hover_color=self.COL["DANGER_HOVER"],
                                      font=(self.UI_FONT, self.UI_SIZE), height=42, corner_radius=14)
        self.quit_btn.pack(side="right", padx=18, pady=14)

        # console
        self.run_console = ctk.CTkFrame(parent, fg_color=self.COL["CARD"], corner_radius=16)
        self.run_console.pack(fill="both", expand=True, padx=4, pady=(6, 10))
        self.pbar = ctk.CTkProgressBar(self.run_console, height=10, corner_radius=8, progress_color=self.COL["ACCENT"])
        self.pbar.pack(fill="x", padx=18, pady=(16, 8))
        self.pbar.set(0.0)

        self.log = ctk.CTkTextbox(self.run_console, corner_radius=12, font=(self.MONO_FONT, self.UI_SIZE),
                                  fg_color=self.COL["CONSOLE_BG"], text_color=self.COL["CONSOLE_FG"])
        self.log.pack(fill="both", expand=True, padx=18, pady=(6, 18))
        self.log.configure(state="disabled")

    def _relayout_sheet_checks(self, event=None):
        """Lay out sheet checkboxes in as many columns as fit, smaller gaps, no right whitespace."""
        # Clear any existing column weights
        try:
            # try a few columns in case the count changed
            for c in range(10):
                self.checks_grid.grid_columnconfigure(c, weight=0)
        except Exception:
            pass

        w = max(1, self.checks_grid.winfo_width())
        # tighter min width and padding to fit more on one row
        min_cb_w = 110
        pad = 12

        # how many columns we can fit
        cols = max(1, (w + pad) // (min_cb_w + pad))

        # give each used column equal weight so the row stretches to the full width
        for c in range(cols):
            self.checks_grid.grid_columnconfigure(c, weight=1, uniform="sheets")

        # place items
        for i, cb in enumerate(self._sheet_checks):
            r, c = divmod(i, cols)
            # sticky="w" keeps them left-aligned, but columns expand so there is no right-side dead zone
            cb.grid(row=r, column=c, sticky="w", padx=8, pady=4)

    # -- islands tab
    def _build_islands_tab(self, parent: ctk.CTkFrame):
        # top controls
        self.island_controls = ctk.CTkFrame(parent, fg_color=self.COL["CARD"], corner_radius=16)
        self.island_controls.pack(fill="x", padx=4, pady=(6, 10))

        left = ctk.CTkFrame(self.island_controls, fg_color=self.COL["CARD"])
        left.pack(side="left", padx=18, pady=12)
        ctk.CTkLabel(left, text="Click a row to set the Active Island.", font=(self.UI_FONT, self.UI_SIZE, "bold"), text_color=self.COL["TEXT"]).pack(anchor="w")
        self.active_island_label = ctk.CTkLabel(left, text="Active island: (none)", font=(self.UI_FONT, self.UI_SIZE), text_color=self.COL["MUTED"])
        self.active_island_label.pack(anchor="w", pady=(6, 0))

        right = ctk.CTkFrame(self.island_controls, fg_color=self.COL["CARD"])
        right.pack(side="right", padx=18, pady=12)
        self.btn_analyze = ctk.CTkButton(right, text="Analyze Islands", command=self._analyze_islands_only,
                                         fg_color=self.COL["ACCENT"], hover_color=self.COL["ACCENT_HOVER"],
                                         font=(self.UI_FONT, self.UI_SIZE), height=36, corner_radius=12)
        self.btn_analyze.pack(side="left", padx=(0, 8))
        self.btn_island_reset = ctk.CTkButton(right, text="Reset", command=self._reset_active_island,
                                              fg_color=self.COL["ACCENT_SOFT"], hover_color=self.COL["ACCENT_SOFT_HOVER"], text_color=self.COL["ACCENT"],
                                              font=(self.UI_FONT, self.UI_SIZE), height=36, corner_radius=12)
        self.btn_island_reset.pack(side="left")

        # adjustable split: table over (bus list + summary)
        self.island_outer = ctk.CTkFrame(parent, fg_color=self.COL["CARD"], corner_radius=16)
        self.island_outer.pack(fill="both", expand=True, padx=4, pady=(6, 10))

        # ttk container styles to avoid light patches in dark mode
        style = ttk.Style()
        try:
            style.configure("Card.TFrame", background=self.COL["CARD"])
            style.configure("Card.TPanedwindow", background=self.COL["CARD"]) 
        except Exception:
            pass

        vpane = ttk.Panedwindow(self.island_outer, orient="vertical", style="Card.TPanedwindow")
        vpane.pack(fill="both", expand=True, padx=8, pady=8)

        top = ttk.Frame(vpane, style="Card.TFrame")
        bottom = ttk.Frame(vpane, style="Card.TFrame")
        vpane.add(top, weight=3)
        vpane.add(bottom, weight=2)

        # top: islands table
        self.tree = ttk.Treeview(
            top,
            columns=("island", "slack", "bus_count", "good"),
            show="headings",
            height=10,
            selectmode="browse",
        )
        self.tree.heading("island", text="Island ID")
        self.tree.heading("slack", text="Slack Bus")
        self.tree.heading("bus_count", text="Bus Count")
        self.tree.heading("good", text="Good Island")
        self.tree.column("island", width=120, anchor="center")
        self.tree.column("slack", width=520, anchor="w")
        self.tree.column("bus_count", width=120, anchor="center")
        self.tree.column("good", width=120, anchor="center")
        self.tree.pack(fill="both", expand=True)
        # click selection → set active island
        self.tree.bind("<ButtonRelease-1>", self._on_island_click)

        # bottom: horizontal split (bus list | summary)
        hpane = ttk.Panedwindow(bottom, orient="horizontal", style="Card.TPanedwindow")
        hpane.pack(fill="both", expand=True)

        buses_frame = ttk.Frame(hpane, style="Card.TFrame")
        summary_frame = ttk.Frame(hpane, style="Card.TFrame")
        hpane.add(buses_frame, weight=1)
        hpane.add(summary_frame, weight=1)

        # bus list
        ctk.CTkLabel(buses_frame, text="Buses in selected island:", font=(self.UI_FONT, self.UI_SIZE, "bold"), text_color=self.COL["TEXT"], fg_color=self.COL["CARD"]).pack(anchor="w", padx=8, pady=(6, 0))
        self.bus_list = ctk.CTkTextbox(buses_frame, corner_radius=10, font=(self.MONO_FONT, self.UI_SIZE),
                                       fg_color=self.COL["CONSOLE_BG"], text_color=self.COL["CONSOLE_FG"])
        self.bus_list.pack(fill="both", expand=True, padx=8, pady=8)
        self.bus_list.configure(state="disabled")

        # summary console
        ctk.CTkLabel(summary_frame, text="Island summary:", font=(self.UI_FONT, self.UI_SIZE, "bold"), text_color=self.COL["TEXT"], fg_color=self.COL["CARD"]).pack(anchor="w", padx=8, pady=(6, 0))
        self.island_summary = ctk.CTkTextbox(summary_frame, corner_radius=10, font=(self.MONO_FONT, self.UI_SIZE),
                                             fg_color=self.COL["CONSOLE_BG"], text_color=self.COL["CONSOLE_FG"])
        self.island_summary.pack(fill="both", expand=True, padx=8, pady=8)
        self.island_summary.configure(state="disabled")

    # -- theme switching
    def _on_theme_change(self, value: str):
        mode = (value or "").lower()
        if mode not in ("light", "dark"):
            return
        self._apply_theme(mode)
        save_conf({"theme": mode})

    def _on_theme_switch(self):
        mode = "dark" if int(self.theme_bool.get() or 0) == 1 else "light"
        self._apply_theme(mode)
        save_conf({"theme": mode})

    def _apply_theme(self, mode: str):
        # Update global appearance
        ctk.set_appearance_mode(mode)
        # Refresh palette
        _, _, _, _, self.COL, _ = setup_appearance(mode)

        # Root + header
        try:
            self.configure(fg_color=self.COL["BG"])
            self.header.configure(fg_color=self.COL["BG"])
            if hasattr(self, "theme_switch"):
                try:
                    self.theme_switch.configure(progress_color=self.COL["ACCENT"]) 
                    self.sun_lbl.configure(text_color=self.COL["TEXT"]) 
                    self.moon_lbl.configure(text_color=self.COL["TEXT"]) 
                    # refresh icons to match text color
                    try:
                        from PIL import Image, ImageDraw, ImageTk  # type: ignore
                        def _sun(color: str):
                            W=18; im=Image.new("RGBA",(W,W),(0,0,0,0)); d=ImageDraw.Draw(im)
                            d.ellipse((4,4,14,14), fill=color)
                            for a in range(0,360,45):
                                d.pieslice((1,1,17,17), a-2, a+2, fill=color)
                            return ImageTk.PhotoImage(im)
                        def _moon(color: str):
                            W=18; im=Image.new("RGBA",(W,W),(0,0,0,0)); d=ImageDraw.Draw(im)
                            d.ellipse((3,3,15,15), fill=color); d.ellipse((7,3,17,15), fill=(0,0,0,0))
                            return ImageTk.PhotoImage(im)
                        self.sun_icon = _sun(self.COL["TEXT"])
                        self.moon_icon = _moon(self.COL["TEXT"])
                        self.sun_lbl.configure(image=self.sun_icon)
                        self.moon_lbl.configure(image=self.moon_icon)
                    except Exception:
                        pass
                except Exception:
                    pass
            # brand holder + image bg
            if hasattr(self, "brand_holder"):
                try:
                    self.brand_holder.configure(bg=self.COL["BG"])
                    for ch in self.brand_holder.winfo_children():
                        try:
                            ch.configure(bg=self.COL["BG"])
                        except Exception:
                            pass
                    # swap brand image for current mode
                    try:
                        logo_png = resource_path("icons", f"cyme_logo_{mode}.png")
                        if not logo_png.exists():
                            logo_png = resource_path("icons", "cyme_logo_light.png")
                        if logo_png.exists() and hasattr(self, "brand_img_label"):
                            new_img = tk.PhotoImage(file=str(logo_png))
                            self._brand_img_ref = new_img
                            self.brand_img_label.configure(image=new_img)
                    except Exception:
                        pass
                except Exception:
                    pass
        except Exception:
            pass

        # Tabs
        try:
            self.tabs.configure(fg_color=self.COL["BG"],
                                segmented_button_selected_color=self.COL["ACCENT"],
                                segmented_button_selected_hover_color=self.COL["ACCENT_HOVER"])
            try:
                seg = self.tabs._segmented_button
                seg.configure(font=(self.UI_FONT, self.UI_SIZE + 2), height=36, corner_radius=12)
            except Exception:
                pass
        except Exception:
            pass

        # Run tab containers
        for f in (getattr(self, "run_card_file", None), getattr(self, "checks_card", None), getattr(self, "run_actions", None), getattr(self, "run_console", None)):
            if f is not None:
                try:
                    f.configure(fg_color=self.COL["CARD"])
                except Exception:
                    pass

        # Buttons in row1/row2 (by discovery to avoid strict references)
        try:
            # row1: label, entry, [browse], [open folder]
            for w in getattr(self, "run_row1", None).winfo_children():
                if isinstance(w, ctk.CTkButton):
                    txt = (w.cget("text") or "").lower()
                    if "browse" in txt:
                        w.configure(fg_color=self.COL["ACCENT"], hover_color=self.COL["ACCENT_HOVER"], text_color=None)
                    else:
                        w.configure(fg_color=self.COL["ACCENT_SOFT"], hover_color=self.COL["ACCENT_SOFT_HOVER"], text_color=self.COL["ACCENT"])
        except Exception:
            pass
        try:
            # row2: label, entry, [save as], [open folder]
            for w in getattr(self, "run_row2", None).winfo_children():
                if isinstance(w, ctk.CTkButton):
                    txt = (w.cget("text") or "").lower()
                    if "save" in txt:
                        w.configure(fg_color=self.COL["ACCENT"], hover_color=self.COL["ACCENT_HOVER"], text_color=None)
                    else:
                        w.configure(fg_color=self.COL["ACCENT_SOFT"], hover_color=self.COL["ACCENT_SOFT_HOVER"], text_color=self.COL["ACCENT"])
        except Exception:
            pass

        # Primary and danger
        try:
            self.run_btn.configure(fg_color=self.COL["ACCENT"], hover_color=self.COL["ACCENT_HOVER"])
            self.quit_btn.configure(fg_color=self.COL["DANGER"], hover_color=self.COL["DANGER_HOVER"])
        except Exception:
            pass

        # Entries styling (file selections)
        try:
            self.in_entry.configure(fg_color=self.COL["INPUT_BG"], border_color=self.COL["INPUT_BORDER"], text_color=self.COL["INPUT_FG"]) 
            self.out_entry.configure(fg_color=self.COL["INPUT_BG"], border_color=self.COL["INPUT_BORDER"], text_color=self.COL["INPUT_FG"]) 
        except Exception:
            pass

        # Checkboxes accents
        try:
            for cb in getattr(self, "_sheet_checks", []):
                cb.configure(fg_color=self.COL["ACCENT"], hover_color=self.COL["ACCENT_HOVER"], text_color=self.COL["TEXT"], border_color=self.COL["BORDER"])
        except Exception:
            pass

        # Console + pbar
        try:
            self.pbar.configure(progress_color=self.COL["ACCENT"], fg_color=self.COL["BORDER"])  # track color
            self.log.configure(fg_color=self.COL["CONSOLE_BG"], text_color=self.COL["CONSOLE_FG"])
        except Exception:
            pass

        # Islands tab frames + controls
        for f in (getattr(self, "island_controls", None), getattr(self, "island_outer", None)):
            if f is not None:
                try:
                    f.configure(fg_color=self.COL["CARD"])
                except Exception:
                    pass
        try:
            self.btn_analyze.configure(fg_color=self.COL["ACCENT"], hover_color=self.COL["ACCENT_HOVER"])
            self.btn_island_reset.configure(fg_color=self.COL["ACCENT_SOFT"], hover_color=self.COL["ACCENT_SOFT_HOVER"], text_color=self.COL["ACCENT"])
            self.bus_list.configure(fg_color=self.COL["CONSOLE_BG"], text_color=self.COL["CONSOLE_FG"])
            self.island_summary.configure(fg_color=self.COL["CONSOLE_BG"], text_color=self.COL["CONSOLE_FG"])
            self.active_island_label.configure(text_color=self.COL["MUTED"])
        except Exception:
            pass

        # ttk look (Treeview)
        try:
            style = ttk.Style()
            try:
                style.theme_use("clam")
            except Exception:
                pass
            # Containers
            try:
                style.configure("Card.TFrame", background=self.COL["CARD"])
                style.configure("Card.TPanedwindow", background=self.COL["CARD"]) 
                style.configure("TFrame", background=self.COL["CARD"]) 
                style.configure("TPanedwindow", background=self.COL["CARD"]) 
            except Exception:
                pass
            style.configure("Treeview",
                            background=self.COL["CARD"],
                            fieldbackground=self.COL["CARD"],
                            foreground=self.COL["TEXT"],
                            bordercolor=self.COL["BORDER"],
                            rowheight=24)
            style.configure("Treeview.Heading",
                            background=self.COL["BG"],
                            foreground=self.COL["TEXT"],
                            bordercolor=self.COL["BORDER"])
            style.map("Treeview",
                      background=[("selected", self.COL["ACCENT_SOFT"])],
                      foreground=[("selected", self.COL["TEXT"])])
        except Exception:
            pass

    # ----- helpers (file picks, folders) --------------------------------------
    def _browse_in(self):
        path = filedialog.askopenfilename(title="Select CYME export",
                                          filetypes=[("CYME export", "*.txt *.xml *.sxst"), ("All files", "*.*")])
        if path:
            self.in_path.set(path)

    def _browse_out(self):
        path = filedialog.asksaveasfilename(title="Save Excel as", defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if path:
            self.out_path.set(path)

    def _open_in_dir(self):
        inp = Path(self.in_path.get() or "")
        folder = (inp.parent if inp else Path.cwd())
        open_in_file_explorer(folder)

    def _open_out_dir(self):
        outp = Path(self.out_path.get() or "")
        folder = (outp.parent if outp else Path.cwd())
        open_in_file_explorer(folder)

    # ----- run pipeline -------------------------------------------------------
    def _start_run(self):
        in_path = Path(self.in_path.get()).expanduser()
        out_path = Path(self.out_path.get()).expanduser()
        if not in_path.exists():
            messagebox.showerror(APP_NAME, "Input file not found.")
            return
        if out_path.suffix.lower() != ".xlsx":
            out_path = out_path.with_suffix(".xlsx")
            self.out_path.set(str(out_path))

        save_conf({"last_input": str(in_path), "last_output": str(out_path)})
        sheets = {name: var.get() for name, var in self.sheet_vars.items()}

        self._set_busy(True); self._clear_log(); self._set_progress(0)

        t = threading.Thread(target=self._run_pipeline_worker, args=(in_path, out_path, sheets), daemon=True)
        t.start()

    def _run_pipeline_worker(self, in_path: Path, out_path: Path, sheets: dict[str, bool]):
        try:
            steps: list[tuple[str, callable]] = []
            if sheets["General"]:        steps.append(("General",        write_general_sheet))
            if sheets["Pins"]:           steps.append(("Pins",           write_pins_sheet))
            if sheets["Bus"]:            steps.append(("Bus",            write_bus_sheet))
            if sheets["Voltage Source"]: steps.append(("Voltage Source", write_voltage_source_sheet))
            if sheets["Load"]:           steps.append(("Load",           write_load_sheet))
            if sheets["Line"]:           steps.append(("Line",           write_line_sheet))
            if sheets["Transformer"]:    steps.append(("Transformer",    write_transformer_sheet))
            if sheets["Switch"]:         steps.append(("Switch",         write_switch_sheet))
            if sheets["Shunt"]:          steps.append(("Shunt",          write_shunt_sheet))

            total = len(steps) + 1
            cur = 0

            # Analyze first
            self._emit("log", f"[1/{total}] Analyzing islands …")
            analyze_and_set_island_context(in_path, per_island_limit=50)
            # Decide export scope now (UI click does NOT prune; we prune only at run time)
            if self.active_island_id is not None:
                self._emit("log", f" → Exporting ONLY island {self.active_island_id}")
                _filter_context_to_island(self, self.active_island_id)
            else:
                self._emit("log", " → No active island chosen: keeping all islands WITH voltage sources; pruning islands without sources")
                _keep_sourceful_islands_context(self, in_path)

            # If user already chose an active island earlier, re-apply it after analysis
            if self.active_island_id is not None:
                try:
                    self._apply_selected_island_context(self.active_island_id)
                    self._emit("log", f" → Active island preserved: {self.active_island_id}")
                except Exception as e:
                    self._emit("log", f" ! Failed to re-apply active island: {e}")

            # Refresh island UI
            self._emit("islands", None)
            cur += 1
            self._emit("progress", int(cur / total * 100))

            # Write workbook
            self._emit("log", f"[2/{total}] Writing workbook → {out_path}")
            out_path.parent.mkdir(parents=True, exist_ok=True)
            with pd.ExcelWriter(out_path, engine="xlsxwriter") as xw:
                for name, fn in steps:
                    self._emit("log", f"  - {name}")
                    fn(xw, in_path)
                    cur += 1
                    self._emit("progress", int(cur / total * 100))

            self._emit("log", f"✔ Done. Wrote: {out_path}")
            self._emit("done", str(out_path))
        except Exception as e:
            self._emit("error", "".join(traceback.format_exception(e)))

    # ----- Analyze-only for Islands tab --------------------------------------
    def _analyze_islands_only(self):
        in_path = Path(self.in_path.get() or "").expanduser()
        if not in_path.exists():
            messagebox.showwarning(APP_NAME, "Select a CYME file on the Run tab first.")
            return
        self._set_busy(True)
        def _worker():
            try:
                analyze_and_set_island_context(in_path, per_island_limit=50)
                _keep_sourceful_islands_context(self, in_path)
                # re-apply active island if one already set
                if self.active_island_id is not None:
                    self._apply_selected_island_context(self.active_island_id)
                self._emit("islands", None)
            except Exception as e:
                self._emit("error", "".join(traceback.format_exception(e)))
            finally:
                self._set_busy(False)
        threading.Thread(target=_worker, daemon=True).start()

    def _on_island_click(self, event):
        if self._handling_island_click:
            return
        row = self.tree.identify_row(event.y)
        if not row:
            return
        try:
            self._handling_island_click = True
            # visually select/focus the clicked row
            self.tree.selection_set(row)
            self.tree.focus(row)

            values = self.tree.item(row).get("values", [])
            if not values:
                return
            island_id = int(values[0])

            # Only record choice & refresh UI — DO NOT prune context here
            self.active_island_id = island_id
            self.active_island_label.configure(text=f"Active island: {island_id}")
            # update the right-hand panels (bus list/summary) without mutating context
            self._refresh_islands_tab(select_island=island_id)
            # (no messagebox here to keep UX smooth)
        finally:
            self._handling_island_click = False

    def _reset_active_island(self):
        in_path = Path(self.in_path.get() or "").expanduser()
        if not in_path.exists():
            messagebox.showwarning(APP_NAME, "Select a CYME file on the Run tab first.")
            return
        analyze_and_set_island_context(in_path, per_island_limit=50)
        self.active_island_id = None
        self._suppress_island_event = False
        self._handling_island_click: bool = False
        self.active_island_label.configure(text="Active island: (none)")
        self._refresh_islands_tab()
        messagebox.showinfo(APP_NAME, "Island selection cleared.")

    # Core: modify island context so only one island is kept (one slack)
    def _apply_selected_island_context(self, active_island: int):
        """
        Keep ONLY the selected island. Do NOT inherit previous bad_buses,
        so a user-chosen island with no source is still exported.
        """
        from Modules.General import get_island_context, set_island_context  # safe import

        ctx = get_island_context() or {}
        bus_to_island: dict[str, int] = dict(ctx.get("bus_to_island", {}))
        islands: dict[int, set[str]] = dict(ctx.get("islands", {}))
        slack_per_island: dict[int, str] = dict(ctx.get("slack_per_island", {}))

        # If Islands dict is not present (older contexts), synthesize it once.
        if not islands:
            islands = {}
            for b, isl in bus_to_island.items():
                islands.setdefault(isl, set()).add(b)

        keep = islands.get(active_island, set())
        # Rebuild bus_to_island to ONLY the selected island
        new_bus_to_island = {b: active_island for b in keep}

        # All buses NOT in the selected island are now "bad"
        bad_buses = set().union(*(v for k, v in islands.items() if k != active_island)) if islands else set()

        new_ctx = {
            "bus_to_island": new_bus_to_island,
            "bad_buses": bad_buses,                      # selected island never appears here
            "slack_per_island": {                        # preserve slack if it was known for this island
                active_island: slack_per_island.get(active_island, "")
            },
            "islands": {active_island: keep},            # optional but tidy
        }
        set_island_context(new_ctx)

    # ----- thread → UI bridge -------------------------------------------------
    def _emit(self, kind: str, payload: object):
        self.events.put((kind, payload))

    def _poll_events(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log":
                    self._append_log(str(payload))
                elif kind == "progress":
                    self._set_progress(int(payload))
                elif kind == "done":
                    self._set_busy(False); messagebox.showinfo(APP_NAME, f"Export complete:\n{payload}")
                elif kind == "error":
                    self._set_busy(False); self._append_log(str(payload)); messagebox.showerror(APP_NAME, "An error occurred.\n\nSee log for details.")
                elif kind == "islands":
                    self._refresh_islands_tab()
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    # ----- UI helpers ---------------------------------------------------------
    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        try:
            self.in_entry.configure(state=state)
            self.out_entry.configure(state=state)
            self.run_btn.configure(state=state)
        except Exception:
            pass

    def _append_log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal"); self.log.delete("1.0", "end"); self.log.configure(state="disabled")

    def _set_progress(self, percent: int):
        p = max(0.0, min(1.0, percent / 100.0))
        self.pbar.set(p); self.update_idletasks()

    # ----- Island UI refresh --------------------------------------------------
    def _refresh_islands_tab(self, select_island: int | None = None):
        # refresh table + labels + summary + bus list
        for i in self.tree.get_children():
            self.tree.delete(i)

        ctx = get_island_context() or {}
        bus_to_island: dict[str, int] = dict(ctx.get("bus_to_island", {}))
        slack_per_island: dict[int, str] = dict(ctx.get("slack_per_island", {}))
        bad_buses: set[str] = set(ctx.get("bad_buses", set()))

        counts: dict[int, int] = {}
        bad_islands: set[int] = set()
        for bus, isl in bus_to_island.items():
            counts[isl] = counts.get(isl, 0) + 1
            if bus in bad_buses:
                bad_islands.add(isl)

        islands = sorted(counts.keys())
        item_id_by_island: dict[int, str] = {}
        for isl in islands:
            slack = slack_per_island.get(isl, "")
            is_good = "Yes" if isl not in bad_islands else "No"
            iid = self.tree.insert("", "end", values=(isl, slack, counts[isl], is_good))
            item_id_by_island[isl] = iid

        # update active label
        if self.active_island_id is not None:
            self.active_island_label.configure(text=f"Active island: {self.active_island_id}")
        else:
            self.active_island_label.configure(text="Active island: (none)")

        # try to select active or requested island
        to_select = select_island or self.active_island_id
        if to_select in item_id_by_island:
            self._suppress_island_event = True
            try:
                self.tree.selection_set(item_id_by_island[to_select])
                self.tree.see(item_id_by_island[to_select])
            finally:
                self._suppress_island_event = False
                self._handling_island_click: bool = False

        # summary
        total_islands = len(islands)
        good_count = total_islands - len(bad_islands)
        lines = [
            f"Islands found: {total_islands}",
            f"Good islands (with source): {good_count}",
            f"Without source: {len(bad_islands)}",
            "",
            "Slack per island:",
        ]
        for isl in islands:
            lines.append(
                f"  - Island {isl}: {slack_per_island.get(isl, '(none)')}  |  Buses: {counts[isl]}  |  Good: {'Yes' if isl not in bad_islands else 'No'}"
            )
        self.island_summary.configure(state="normal")
        self.island_summary.delete("1.0", "end")
        self.island_summary.insert("end", "\n".join(lines))
        self.island_summary.configure(state="disabled")

        # bus list for current selection (if any), else active, else placeholder
        sel = self.tree.selection()
        if sel:
            item = self.tree.item(sel[0]); vals = item.get("values", [])
            isl = int(vals[0]) if vals else None
        else:
            isl = self.active_island_id

        self._populate_bus_list_for_island(isl, bus_to_island)

    def _populate_bus_list_for_island(self, isl: int | None, bus_to_island: dict[str, int]):
        self.bus_list.configure(state="normal")
        self.bus_list.delete("1.0", "end")
        if isl is None:
            self.bus_list.insert("end", "(select an island to view its buses)")
            self.bus_list.configure(state="disabled")
            return
        buses = sorted([b for b, i in bus_to_island.items() if i == isl])
        if not buses:
            self.bus_list.insert("end", "(no buses in this island)")
        else:
            self.bus_list.insert("end", "\n".join(buses))
        self.bus_list.configure(state="disabled")

# ---- Entrypoint --------------------------------------------------------------
def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
