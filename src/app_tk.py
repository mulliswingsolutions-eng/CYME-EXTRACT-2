# src/app_tk.py  - Premium UI + Island workflow (click row to activate)
from __future__ import annotations
import os, sys, json, threading, queue, traceback, platform, subprocess
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, TYPE_CHECKING, cast

import customtkinter as ctk
import pandas as pd

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

APP_NAME = "CYME ? XLSX Extractor"
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

        # Use a resampling method compatible across Pillow versions (no direct BICUBIC reference)
        _Resampling = getattr(Image, "Resampling", None)
        _filter = getattr(_Resampling, "LANCZOS", getattr(Image, "LANCZOS", 1))
        out = img.resize((256, 256), _filter)
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
        self.events: "queue.Queue[tuple[str, Any]]" = queue.Queue()
        self.in_path = tk.StringVar(value=conf.get("last_input", ""))
        self.out_path = tk.StringVar(value=conf.get("last_output", str(Path.cwd() / "CYME_Extract.xlsx")))
        self.sheet_vars: dict[str, tk.BooleanVar] = {name: tk.BooleanVar(value=DEFAULT_SHEETS[name]) for name in DEFAULT_SHEETS}
        self.active_island_id: int | None = None  # click-to-activate
        self._suppress_island_event: bool = False  # guard to prevent feedback loops
        self._handling_island_click: bool = False

        # UI
        self._build_header()
        self._build_tabs()
        # (No dark mode - single palette already applied)
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
            from PIL import Image, ImageDraw  # type: ignore
            def sun_img(color: str) -> ctk.CTkImage:
                W = 18
                im = Image.new("RGBA", (W, W), (0,0,0,0))
                d = ImageDraw.Draw(im)
                # core
                d.ellipse((4,4,14,14), fill=color)
                # rays
                for a in range(0,360,45):
                    # tiny rectangles as rays
                    d.pieslice((1,1,17,17), a-2, a+2, fill=color)
                return ctk.CTkImage(light_image=im, size=(W, W))
            def moon_img(color: str) -> ctk.CTkImage:
                W = 18
                im = Image.new("RGBA", (W, W), (0,0,0,0))
                d = ImageDraw.Draw(im)
                d.ellipse((3,3,15,15), fill=color)
                d.ellipse((7,3,17,15), fill=(0,0,0,0))
                return ctk.CTkImage(light_image=im, size=(W, W))
            self.sun_icon = sun_img(self.COL["TEXT"])
            self.moon_icon = moon_img(self.COL["TEXT"])
        except Exception:
            pass

        row = ctk.CTkFrame(parent, fg_color=self.COL["BG"])
        row.pack(side="right")
        self.sun_lbl = ctk.CTkLabel(row, text="", image=self.sun_icon, text_color=self.COL["TEXT"]) if self.sun_icon else ctk.CTkLabel(row, text="?", text_color=self.COL["TEXT"]) 
        self.sun_lbl.pack(side="left", padx=(0,6))
        self.theme_bool = tk.IntVar(value=1 if str(ctk.get_appearance_mode()).lower() == "dark" else 0)
        self.theme_switch = ctk.CTkSwitch(row, text="", command=self._on_theme_switch,
                                          variable=self.theme_bool, onvalue=1, offvalue=0,
                                          width=52, height=28, progress_color=self.COL["ACCENT"])
        self.theme_switch.pack(side="left")
        self.moon_lbl = ctk.CTkLabel(row, text="", image=self.moon_icon, text_color=self.COL["TEXT"]) if self.moon_icon else ctk.CTkLabel(row, text="??", text_color=self.COL["TEXT"]) 
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
        ctk.CTkButton(row1, text="Browse...", command=self._browse_in,
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
        ctk.CTkButton(row2, text="Save As...", command=self._browse_out,
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
        self.btn_island_reset.pack(side="left", padx=(0, 8))

        # Unused objects policy: Comment vs Remove
        ctk.CTkLabel(right, text="Unused objects:", font=(self.UI_FONT, self.UI_SIZE), text_color=self.COL["TEXT"]).pack(side="left", padx=(8, 6))
        self.prune_mode = tk.StringVar(value="Comment")
        self.prune_menu = ctk.CTkOptionMenu(right, values=["Comment", "Remove"], variable=self.prune_mode,
                                            fg_color=self.COL["ACCENT_SOFT"], button_color=self.COL["ACCENT"],
                                            button_hover_color=self.COL["ACCENT_HOVER"], text_color=self.COL["TEXT"],
                                            font=(self.UI_FONT, self.UI_SIZE))
        self.prune_menu.pack(side="left")

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
        # click selection ? set active island
        self.tree.bind("<ButtonRelease-1>", self._on_island_click)

        # bottom: horizontal split (bus list | map)
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

        # island map canvas
        ctk.CTkLabel(summary_frame, text="Island map:", font=(self.UI_FONT, self.UI_SIZE, "bold"), text_color=self.COL["TEXT"], fg_color=self.COL["CARD"]).pack(anchor="w", padx=8, pady=(6, 0))
        self.map_container = summary_frame
        self.island_map = tk.Canvas(summary_frame, highlightthickness=0, background=self.COL["CARD"])
        self.island_map.pack(fill="both", expand=True, padx=8, pady=8)
        # Loading overlay (indeterminate progress) for async map builds
        # Queues/worker state for async map rendering
        self._map_queue: "queue.Queue[tuple[int, int | None, dict]]" = queue.Queue()
        self._map_job_counter: int = 0
        # Circular spinner overlay (themed)
        try:
            self.map_spinner = tk.Canvas(self.island_map, width=36, height=36, highlightthickness=0, bg=self.COL["CARD"])
            self._spinner_angle = 0
            self._spinner_running = False
        except Exception:
            self.map_spinner = None
        # Map view state for auto-fit/center
        self._map_fitted: bool = False
        self._last_map_island: int | None = None
        # Redraw on resize (and trigger refit)
        self.island_map.bind("<Configure>", self._on_map_configure)
        # Zoom with mouse wheel (Windows/macOS) and Button-4/5 (X11)
        self.island_map.bind("<MouseWheel>", self._on_map_wheel)
        self.island_map.bind("<Button-4>", lambda e: self._on_map_wheel(e, delta=120))
        self.island_map.bind("<Button-5>", lambda e: self._on_map_wheel(e, delta=-120))
        # Pan by dragging with left mouse button
        self.island_map.bind("<ButtonPress-1>", lambda e: self.island_map.scan_mark(e.x, e.y))
        self.island_map.bind("<B1-Motion>", lambda e: self.island_map.scan_dragto(e.x, e.y, gain=1))

    def _show_map_loader(self, show: bool):
        try:
            if not hasattr(self, 'map_spinner') or self.map_spinner is None:
                return
            if show:
                w = max(50, int(self.island_map.winfo_width() or 0))
                h = max(50, int(self.island_map.winfo_height() or 0))
                self.map_spinner.configure(bg=self.COL.get("CARD", "#FFFFFF"))
                self.map_spinner.place(x=w//2, y=h//2, anchor='center')
                self._spinner_running = True
                self._spinner_tick()
            else:
                self._spinner_running = False
                self.map_spinner.place_forget()
        except Exception:
            pass

    def _spinner_tick(self):
        try:
            if not getattr(self, '_spinner_running', False) or self.map_spinner is None:
                return
            c = self.map_spinner
            c.delete('all')
            W = int(c.winfo_width() or 36); H = int(c.winfo_height() or 36)
            s = min(W, H)
            pad = 4
            x0, y0, x1, y1 = pad, pad, s - pad, s - pad
            # base ring
            c.create_oval(x0, y0, x1, y1, outline=self.COL.get('BORDER', '#E5E7EB'), width=2)
            # moving arc
            ang = int(getattr(self, '_spinner_angle', 0)) % 360
            extent = 270
            c.create_arc(x0, y0, x1, y1, start=ang, extent=extent, style='arc', outline=self.COL.get('ACCENT', '#A78BFA'), width=3)
            self._spinner_angle = (ang + 18) % 360
            self.after(60, self._spinner_tick)
        except Exception:
            pass

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
                        from PIL import Image, ImageDraw  # type: ignore
                        def _sun(color: str):
                            W=18; im=Image.new("RGBA",(W,W),(0,0,0,0)); d=ImageDraw.Draw(im)
                            d.ellipse((4,4,14,14), fill=color)
                            for a in range(0,360,45):
                                d.pieslice((1,1,17,17), a-2, a+2, fill=color)
                            return ctk.CTkImage(light_image=im, size=(W, W))
                        def _moon(color: str):
                            W=18; im=Image.new("RGBA",(W,W),(0,0,0,0)); d=ImageDraw.Draw(im)
                            d.ellipse((3,3,15,15), fill=color); d.ellipse((7,3,17,15), fill=(0,0,0,0))
                            return ctk.CTkImage(light_image=im, size=(W, W))
                        self.sun_icon = _sun(self.COL["TEXT"])
                        self.moon_icon = _moon(self.COL["TEXT"])
                        self.sun_lbl.configure(image=self.sun_icon)
                        self.moon_lbl.configure(image=self.moon_icon)
                    except Exception:
                        pass
                except Exception:
                    pass
            # brand image swap to match theme (avoid unknown attributes)
            try:
                logo_png = resource_path("icons", f"cyme_logo_{mode}.png")
                if not logo_png.exists():
                    logo_png = resource_path("icons", "cyme_logo_light.png")
                if logo_png.exists() and hasattr(self, "brand_img_label"):
                    new_img = ctk.CTkImage(light_image=Image.open(str(logo_png)), size=(28, 28))  # type: ignore[name-defined]
                    self.brand_img = new_img  # keep reference
                    self.brand_img_label.configure(image=new_img)
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
        # Buttons in row1/row2 (by discovery to avoid strict references)
        try:
            # row1: label, entry, [browse], [open folder]
            row1 = getattr(self, "run_row1", None)
            if row1 is not None:
                for w in row1.winfo_children():
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
            row2 = getattr(self, "run_row2", None)
            if row2 is not None:
                for w in row2.winfo_children():
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
            # Canvas background to match card color
            if hasattr(self, 'island_map') and self.island_map is not None:
                try:
                    self.island_map.configure(background=self.COL["CARD"])
                except Exception:
                    pass
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
            steps: list[tuple[str, Callable[[Any, Path], None]]] = []
            if sheets.get("General", False):        steps.append(("General",        write_general_sheet))
            if sheets.get("Pins", False):           steps.append(("Pins",           write_pins_sheet))
            if sheets.get("Bus", False):            steps.append(("Bus",            write_bus_sheet))
            if sheets.get("Voltage Source", False): steps.append(("Voltage Source", write_voltage_source_sheet))
            if sheets.get("Load", False):           steps.append(("Load",           write_load_sheet))
            if sheets.get("Line", False):           steps.append(("Line",           write_line_sheet))
            if sheets.get("Transformer", False):    steps.append(("Transformer",    write_transformer_sheet))
            if sheets.get("Switch", False):         steps.append(("Switch",         write_switch_sheet))
            if sheets.get("Shunt", False):          steps.append(("Shunt",          write_shunt_sheet))

            total = len(steps) + 1
            cur = 0

            # Analyze first
            self._emit("log", "Analyzing islands")
            analyze_and_set_island_context(in_path, per_island_limit=50)
            # Decide export scope now (UI click does NOT prune; we prune only at run time)
            if self.active_island_id is not None:
                self._emit("log", f" - Export scope: only island {self.active_island_id}")
                _filter_context_to_island(self, self.active_island_id)
            else:
                self._emit("log", " - No island selected. Keeping islands with voltage sources; excluding islands without sources")
                _keep_sourceful_islands_context(self, in_path)

            # If user already chose an active island earlier, re-apply it after analysis
            if self.active_island_id is not None:
                try:
                    self._apply_selected_island_context(self.active_island_id)
                    self._emit("log", f" - Active island preserved: {self.active_island_id}")
                except Exception as e:
                    self._emit("log", f" ! Could not re-apply the active island: {e}")

            # Refresh island UI
            self._emit("islands", None)
            cur += 1
            self._emit("progress", int(cur / total * 100))

            # Write workbook
            self._emit("log", f"Writing workbook - {out_path}")
            out_path.parent.mkdir(parents=True, exist_ok=True)
            opened_path = out_path
            try:
                writer = pd.ExcelWriter(out_path, engine="xlsxwriter")
            except PermissionError:
                from datetime import datetime
                alt = out_path.with_name(f"{out_path.stem}__{datetime.now().strftime('%Y%m%d_%H%M%S')}{out_path.suffix}")
                opened_path = alt
                self._emit("log", f" ! Output file is in use. Writing to new file: {alt}")
                writer = pd.ExcelWriter(alt, engine="xlsxwriter")
            with writer as xw:
                # Apply prune policy to the island context (comment vs remove)
                try:
                    from Modules.General import get_island_context, set_island_context  # lazy import
                    ctx = get_island_context() or {}
                    mode = (self.prune_mode.get() if hasattr(self, 'prune_mode') else 'Comment') or 'Comment'
                    ctx["prune_mode"] = "remove" if str(mode).strip().lower().startswith("remove") else "comment"
                    set_island_context(ctx)
                except Exception:
                    pass
                for name, fn in steps:
                    # Log each sheet succinctly
                    self._emit("log", f"  - {name}")
                    fn(xw, in_path)
                    cur += 1
                    self._emit("progress", int(cur / total * 100))

            self._emit("log", f"Export complete: {opened_path}")
            self._emit("done", str(opened_path))
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

            # removed stray duplicate error emission

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

            # Only record choice & refresh UI - DO NOT prune context here
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

    # ----- thread ? UI bridge -------------------------------------------------
    def _emit(self, kind: str, payload: Any):
        self.events.put((kind, payload))

    def _poll_events(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log":
                    self._append_log(str(payload))
                elif kind == "progress":
                    # payload is expected int, but accept Any and coerce safely
                    try:
                        self._set_progress(int(payload))
                    except Exception:
                        self._set_progress(0)
                elif kind == "done":
                    self._set_busy(False); messagebox.showinfo(APP_NAME, f"Export complete:\n{payload}")
                elif kind == "error":
                    self._set_busy(False); self._append_log(str(payload)); messagebox.showerror(APP_NAME, "An error occurred.\n\nSee log for details.")
                elif kind == "islands":
                    self._refresh_islands_tab()
        except queue.Empty:
            pass
        # Poll async island map build results too
        try:
            while True:
                job_id, isl, data = self._map_queue.get_nowait()
                if job_id != getattr(self, '_map_job_counter', 0):
                    continue  # stale job, ignore
                if isinstance(data, dict) and data.get('error'):
                    self._show_map_loader(False)
                    continue
                # Instrumentation: log quick diagnostics if present
                try:
                    diag = data.get('diag') if isinstance(data, dict) else None
                    if isinstance(diag, dict):
                        a = diag.get('anchors'); ps = diag.get('poly_sections'); pp = diag.get('poly_points'); sn = diag.get('synthetic_nodes')
                        self._append_log(f"Map diag — anchors={a}, poly_sections={ps}, poly_points={pp}, synthetic={sn}")
                        if (a or 0) == 0 and (ps or 0) == 0:
                            self._append_log("No diagram geometry — using fallback layout")
                except Exception:
                    pass
                try:
                    self._render_island_map_from_data(data)
                finally:
                    self._show_map_loader(False)
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

        # If nothing selected/active yet, auto-select a sensible default after analysis
        if not select_island and self.active_island_id is None and item_id_by_island:
            # Prefer a "good" island (not marked bad), else the first by id
            try:
                preferred = next((i for i in islands if i not in bad_islands), None)
            except Exception:
                preferred = None
            select_island = preferred if preferred is not None else islands[0]
            self.active_island_id = select_island

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

        # draw/update map for current selection (async)
        sel = self.tree.selection()
        if sel:
            item = self.tree.item(sel[0]); vals = item.get("values", [])
            isl_for_map = int(vals[0]) if vals else None
        else:
            isl_for_map = self.active_island_id
        self._start_island_map_job(isl_for_map, bus_to_island)

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

    # ----- Island map drawing -------------------------------------------------
    def _redraw_island_map(self):
        try:
            ctx = get_island_context() or {}
            bus_to_island: dict[str, int] = dict(ctx.get("bus_to_island", {}))
            sel = self.tree.selection()
            if sel:
                item = self.tree.item(sel[0]); vals = item.get("values", [])
                isl = int(vals[0]) if vals else None
            else:
                isl = self.active_island_id
            # Reset fit when island changes
            if getattr(self, "_last_map_island", None) != isl:
                try:
                    self._map_fitted = False
                except Exception:
                    pass
                self._last_map_island = isl
            self._start_island_map_job(isl, bus_to_island)
        except Exception:
            pass

    def _on_map_configure(self, event=None):
        try:
            # Any canvas size change should trigger a refit on next draw
            self._map_fitted = False
        except Exception:
            pass
        self._redraw_island_map()

    # Async trigger for island map rendering so UI stays responsive
    def _start_island_map_job(self, isl: int | None, bus_to_island: dict[str, int]):
        try:
            # Cancel previous pending jobs by bumping the counter
            self._map_job_counter = int(getattr(self, '_map_job_counter', 0)) + 1
            job_id = self._map_job_counter
            # Show loader now
            self._show_map_loader(True)
            # Clear previous map immediately so old island disappears right away
            try:
                if hasattr(self, 'island_map') and self.island_map is not None:
                    self.island_map.delete('all')
            except Exception:
                pass

            # Launch compute on a background thread; main thread will render
            # once results arrive via _map_queue.
            in_path = Path(self.in_path.get() or "").expanduser()
            def _worker():
                try:
                    data = self._compute_island_map_data(isl, bus_to_island, in_path)
                except Exception as e:
                    data = {"error": str(e)}
                try:
                    self._map_queue.put((job_id, isl, data))
                except Exception:
                    pass
            threading.Thread(target=_worker, daemon=True).start()
        except Exception:
            # On any failure, fall back to direct draw and hide loader safely
            try:
                data = self._compute_island_map_data(isl, bus_to_island, Path(self.in_path.get() or "").expanduser())
                self._map_queue.put((self._map_job_counter, isl, data))
            finally:
                self._show_map_loader(False)

    # Compatibility layer for async map build/render.
    # Move heavy XML parsing and normalization to a background thread
    # and only do the lightweight canvas mapping/drawing on the UI thread.
    def _compute_island_map_data(self, isl: int | None, bus_to_island: dict[str, int], in_path: Path) -> dict:
        out: dict[str, Any] = {"version": 4, "isl": isl}
        try:
            if isl is None or not isinstance(bus_to_island, dict):
                out["error"] = "no island"
                return out
            # Selected island nodes
            nodes = sorted([b for b, i in bus_to_island.items() if i == isl])
            if not nodes:
                out["coords_unit"] = {}
                out["edges"] = []
                out["polylines_unit"] = {}
                return out

            import xml.etree.ElementTree as ET
            import math as _m
            from Modules.General import safe_name as _safe
            txt = in_path.read_text(encoding='utf-8', errors='ignore')
            xml_root = ET.fromstring(txt)

            nodes_set = set()  # will be set after ID normalization

            # Build adjacency and capture polylines per section
            adj: dict[str, set[str]] = {n: set() for n in nodes}
            edges: list[tuple[str, str]] = []
            section_polylines: dict[str, list[list[tuple[float, float]]]] = {}
            # Track tertiary per base edge to allow polyline splitting later
            tert_by_edge: dict[str, set[str]] = {}
            # Device collections
            bus_sources: set[str] = set()
            bus_loads: dict[str, int] = {}
            bus_shunts: dict[str, int] = {}
            inline_devs: dict[str, list[dict[str, str]]] = {}

            def edge_key(a: str, b: str) -> str:
                return "|".join(sorted((a, b)))

            def _finite(v: float | None) -> bool:
                try:
                    return v is not None and _m.isfinite(v) and abs(float(v)) < 1e12
                except Exception:
                    return False

            def _read_xy(el: ET.Element) -> tuple[float | None, float | None]:
                # Strictly geometry fields only
                cand = [
                    (el.findtext('X'), el.findtext('Y')),
                    (el.findtext('PosX'), el.findtext('PosY')),
                    (el.findtext('CoordX'), el.findtext('CoordY')),
                    (el.findtext('MapX'), el.findtext('MapY')),
                ]
                # Attribute forms
                cand.extend([
                    (el.get('x'), el.get('y')),
                    (el.get('X'), el.get('Y')),
                    (el.get('XCoord'), el.get('YCoord')),
                    (el.get('XCOORD'), el.get('YCOORD')),
                ])
                for xs, ys in cand:
                    try:
                        xv = float(xs) if xs not in (None, '') else None
                        yv = float(ys) if ys not in (None, '') else None
                    except Exception:
                        xv = None; yv = None
                    if _finite(xv) and _finite(yv):
                        return xv, yv
                # Embedded Position node
                pos = el.find('Position') or el.find('Coordinates')
                if pos is not None:
                    px, py = _read_xy(pos)
                    if _finite(px) and _finite(py):
                        return px, py
                return None, None

            def _gather_section_polylines(sec: ET.Element) -> list[list[tuple[float, float]]]:
                outp: list[list[tuple[float, float]]] = []
                # Accept only known geometry containers
                holders = [
                    sec.find('Breakpoints'),
                    sec.find('ShapePoints'),
                    sec.find('Polyline'),
                    sec.find('IntermediatePoints'),  # some templates store the path here
                ]
                for holder in holders:
                    if holder is None:
                        continue
                    pts: list[tuple[float, float]] = []
                    for child in list(holder):
                        tag = child.tag.split('}')[-1].lower()
                        if tag in ('breakpoint', 'point'):
                            x, y = _read_xy(child)
                            if _finite(x) and _finite(y) and not (abs(float(x)) == 0.0 and abs(float(y)) == 0.0):
                                pts.append((float(x), float(y)))  # type: ignore[arg-type]
                    if pts:
                        outp.append(pts)
                return outp

            # ID normalizer wraps project safe_name and adds phase/whitespace handling
            def _norm_id(s: str) -> str:
                try:
                    t = _safe((s or '').strip())
                    t = t.replace('\n', '').replace('\r', '')
                    t = t.replace('__', '-')
                    t = t.casefold()
                    # strip trailing phase suffix _a/_b/_c if present
                    if len(t) > 2 and t.endswith(('_a','_b','_c')):
                        t = t[:-2]
                    return t
                except Exception:
                    return (s or '').strip()

            # Normalize node ids first
            nodes = [(_safe(n) or n) for n in nodes]
            nodes = [(_safe(n) or n) for n in nodes]  # ensure safe then norm
            def _norm_id(s: str) -> str:
                try:
                    t = _safe((s or '').strip())
                    t = t.replace('\n', '').replace('\r', '')
                    t = t.replace('__', '-')
                    t = t.casefold()
                    if len(t) > 2 and t.endswith(('_a','_b','_c')):
                        t = t[:-2]
                    return t
                except Exception:
                    return (s or '').strip()
            nodes = [_norm_id(n) for n in nodes]
            nodes_set = set(nodes)

            for sec in xml_root.findall('.//Sections/Section'):
                fb = _norm_id((sec.findtext('FromNodeID') or ''))
                tb = _norm_id((sec.findtext('ToNodeID') or ''))
                if not fb or not tb:
                    continue
                if fb in nodes_set and tb in nodes_set:
                    if tb not in adj[fb]:
                        adj[fb].add(tb); adj[tb].add(fb)
                        edges.append((fb, tb))
                # Tertiary: split into two logical edges
                tert = _norm_id((sec.findtext('TertiaryNodeID') or ''))
                if tert:
                    if tert in nodes_set and fb in nodes_set:
                        if tert not in adj[fb]:
                            adj[fb].add(tert); adj[tert].add(fb)
                            edges.append((fb, tert))
                    if tert in nodes_set and tb in nodes_set:
                        if tb not in adj[tert]:
                            adj[tert].add(tb); adj[tb].add(tert)
                            edges.append((tert, tb))
                    # Record for later polyline splitting
                    tert_by_edge.setdefault(edge_key(fb, tb), set()).add(tert)

                # Geometry polylines if present
                polylists = _gather_section_polylines(sec)
                if polylists and fb and tb:
                    section_polylines.setdefault(edge_key(fb, tb), []).extend(polylists)

            # Devices under this section
            devs = sec.find('./Devices')
            if devs is not None:
                # Loads attach to From bus per sheet rules
                if devs.find('SpotLoad') is not None or devs.find('DistributedLoad') is not None:
                    if fb in nodes_set:
                        bus_loads[fb] = bus_loads.get(fb, 0) + 1
                # Shunt devices at From bus
                if devs.find('ShuntCapacitor') is not None or devs.find('ShuntReactor') is not None:
                    if fb in nodes_set:
                        bus_shunts[fb] = bus_shunts.get(fb, 0) + 1
                # Switch-like inline devices
                for tag in ('Switch', 'Sectionalizer', 'Breaker', 'Fuse', 'Recloser'):
                    for d in devs.findall(tag):
                        loc = (d.findtext('Location') or 'Middle').strip().lower()
                        name = (d.findtext('Name') or d.findtext('DeviceID') or tag)
                        # heuristic state parse
                        state = (d.findtext('NormalOpen') or d.findtext('NormallyOpen') or d.findtext('Open') or d.findtext('Status') or '')
                        closed: bool | None
                        sv = (state or '').strip().lower()
                        if sv in ('open', '1', 'true'):
                            closed = False
                        elif sv in ('closed', '0', 'false'):
                            closed = True
                        else:
                            closed = None
                        rec = {'type': 'switch', 'loc': loc}
                        if name:
                            rec['name'] = name
                        if closed is not None:
                            rec['closed'] = closed
                        inline_devs.setdefault(edge_key(fb, tb), []).append(rec)
                # Some miscellaneous behave as inline switches (RB, LA)
                for d in devs.findall('Miscellaneous'):
                    did = ((d.findtext('DeviceID') or '').strip().upper())
                    if did in {'RB', 'LA'}:
                        loc = (d.findtext('Location') or 'Middle').strip().lower()
                        inline_devs.setdefault(edge_key(fb, tb), []).append({'type': 'switch', 'loc': loc, 'name': did})
                # Transformers inline (or regulators)
                if devs.find('Transformer') is not None or devs.find('Regulator') is not None:
                    xf = devs.find('Transformer')
                    loc = (xf.findtext('Location') if xf is not None else 'Middle') or 'Middle'
                    name = (xf.findtext('Name') if xf is not None else '') or 'Transformer'
                    inline_devs.setdefault(edge_key(fb, tb), []).append({'type': 'xfmr', 'loc': loc.strip().lower(), 'name': name})

            # Real coordinates per node (best-effort)
            def _num(x: str | None) -> float | None:
                try:
                    return float(x) if x not in (None, '') else None
                except Exception:
                    return None
            coords_real: dict[str, tuple[float, float]] = {}
            used_geo = False  # lat/long detected
            for elem in xml_root.iter():
                tag = elem.tag.split('}')[-1].lower()
                if tag in ('tag', 'tags'):
                    continue  # ignore label containers entirely
                # NodeID may be in text or attribute; try both
                nid_raw = (elem.findtext('NodeID') or elem.get('NodeID') or elem.get('NodeId') or elem.get('node') or '')
                if not nid_raw:
                    continue
                # Prefer diagram units first (X/Y families)
                nx, ny = _read_xy(elem)
                # Else try geographic
                if not (_finite(nx) and _finite(ny)):
                    lon_txt = elem.findtext('Longitude'); lat_txt = elem.findtext('Latitude')
                    try:
                        lon = float(lon_txt) if lon_txt not in (None, '') else None
                        lat = float(lat_txt) if lat_txt not in (None, '') else None
                    except Exception:
                        lon = None; lat = None
                    if _finite(lon) and _finite(lat):
                        # Web Mercator projection to planar meters
                        R = 6378137.0
                        lon_rad = _m.radians(float(lon))
                        lat_rad = _m.radians(max(-85.06, min(85.06, float(lat))))
                        nx = R * lon_rad
                        ny = R * _m.log(_m.tan(_m.pi/4.0 + lat_rad/2.0))
                        used_geo = True
                if not (_finite(nx) and _finite(ny)):
                    continue
                if abs(float(nx)) == 0.0 and abs(float(ny)) == 0.0:
                    continue
                nid = _norm_id(nid_raw)
                if nid in nodes_set:
                    coords_real.setdefault(nid, (float(nx), float(ny)))  # type: ignore[arg-type]

            # Voltage sources from Substation topo (normalize ids)
            try:
                for topo in xml_root.findall('.//Topo'):
                    ntype = (topo.findtext('NetworkType') or '').strip().lower()
                    eq_mode = (topo.findtext('EquivalentMode') or '').strip()
                    if ntype != 'substation' or eq_mode == '1':
                        continue
                    srcs = topo.find('./Sources')
                    if srcs is None:
                        continue
                    for src in srcs.findall('./Source'):
                        nid = _norm_id(src.findtext('SourceNodeID') or '')
                        if nid and nid in nodes_set:
                            bus_sources.add(nid)
            except Exception:
                pass

            # Snap polyline endpoints to node anchors in RAW space; promote anchors from polylines when needed
            try:
                # Compute a robust epsilon based on current raw span
                raw_xs = [x for (x, _) in coords_real.values()]
                raw_ys = [y for (_, y) in coords_real.values()]
                for plist in section_polylines.values():
                    for pts in plist:
                        for (x, y) in pts:
                            raw_xs.append(x); raw_ys.append(y)
                if raw_xs and raw_ys:
                    rx_span = max(1.0, max(raw_xs) - min(raw_xs))
                    ry_span = max(1.0, max(raw_ys) - min(raw_ys))
                    eps = 1e-6 * max(rx_span, ry_span)
                else:
                    eps = 1e-3

                def _snap_poly(poly: list[tuple[float, float]], frm_xy: tuple[float, float] | None, to_xy: tuple[float, float] | None, eps: float) -> list[tuple[float, float]]:
                    if not poly or len(poly) < 2:
                        return poly
                    p = list(poly)
                    if frm_xy is not None:
                        d0 = (p[0][0]-frm_xy[0])**2 + (p[0][1]-frm_xy[1])**2
                        d1 = (p[-1][0]-frm_xy[0])**2 + (p[-1][1]-frm_xy[1])**2
                        if min(d0, d1) <= eps*eps:
                            if d0 <= d1:
                                p[0] = frm_xy
                            else:
                                p[-1] = frm_xy
                    if to_xy is not None:
                        d0 = (p[0][0]-to_xy[0])**2 + (p[0][1]-to_xy[1])**2
                        d1 = (p[-1][0]-to_xy[0])**2 + (p[-1][1]-to_xy[1])**2
                        if min(d0, d1) <= eps*eps:
                            if d1 <= d0:
                                p[-1] = to_xy
                            else:
                                p[0] = to_xy
                    return p

                poly_anchored: set[str] = set()
                for k, plist in list(section_polylines.items()):
                    try:
                        u, v = k.split('|', 1)
                    except Exception:
                        continue
                    u_xy = coords_real.get(u)
                    v_xy = coords_real.get(v)
                    new_list: list[list[tuple[float, float]]] = []
                    for pts in plist:
                        sp = _snap_poly(pts, u_xy, v_xy, eps)
                        new_list.append(sp)
                        if u_xy is None and v_xy is None and sp:
                            coords_real[u] = sp[0]
                            coords_real[v] = sp[-1]
                            poly_anchored.add(u); poly_anchored.add(v)
                            u_xy = coords_real.get(u); v_xy = coords_real.get(v)
                    section_polylines[k] = new_list
            except Exception:
                poly_anchored = set()

            # Fallback BFS layout for structure
            from collections import deque as _dq
            # Try to root at slack if available
            try:
                ctx = get_island_context() or {}
                slack_per_island: dict[int, str] = dict(ctx.get('slack_per_island', {}))
                root_node = slack_per_island.get(isl) or (nodes[0] if nodes else None)
            except Exception:
                root_node = nodes[0] if nodes else None
            depth: dict[str, int] = {}
            if root_node:
                q = _dq([root_node]); depth[root_node] = 0
                while q:
                    u = q.popleft()
                    for v in sorted(adj.get(u, set())):
                        if v not in depth:
                            depth[v] = depth[u] + 1
                            q.append(v)
            for n in nodes:
                depth.setdefault(n, 0)
            columns: dict[int, list[str]] = {}
            for n in nodes:
                columns.setdefault(depth[n], []).append(n)
            for col in columns.values():
                col.sort()
            coords_bfs: dict[str, tuple[float, float]] = {}
            for dlevel, col_nodes in columns.items():
                k = max(1, len(col_nodes))
                for i, n in enumerate(col_nodes):
                    x = float(dlevel)
                    y = float(i) / float(k - 1 if k > 1 else 1)
                    coords_bfs[n] = (x, y)
            try:
                dmax = max((xy[0] for xy in coords_bfs.values()), default=1.0)
                if dmax <= 0:
                    dmax = 1.0
                for n, (x, y) in list(coords_bfs.items()):
                    coords_bfs[n] = (x / dmax, y)
            except Exception:
                pass

            # If there are two or more anchors, align BFS axis with the anchors' principal direction (PCA)
            try:
                anchors_u = [coords_real.get(n) for n in nodes if n in coords_real]
                anchors_u = [(x, y) for (x, y) in anchors_u if isinstance(x, float) and isinstance(y, float)]
                if len(anchors_u) >= 2:
                    import math as _pm
                    ax = [x for (x, _) in anchors_u]; ay = [y for (_, y) in anchors_u]
                    mx = sum(ax)/len(ax); my = sum(ay)/len(ay)
                    vx = sum((x-mx)*(x-mx) for x in ax); vy = sum((y-my)*(y-my) for y in ay); vxy = sum((ax[i]-mx)*(ay[i]-my) for i in range(len(ax)))
                    theta = 0.5 * _pm.atan2(2.0*vxy, (vx - vy) if (vx!=vy or vxy!=0) else 1.0)
                    ct, st = _pm.cos(theta), _pm.sin(theta)
                    coords_bfs = {n: (x*ct - y*st, x*st + y*ct) for n,(x,y) in coords_bfs.items()}
            except Exception:
                pass

            # Collect polyline raw points (for normalization even if no node coords)
            poly_points: list[tuple[float, float]] = []
            for plist in section_polylines.values():
                for pts in plist:
                    poly_points.extend(pts)

            # Build a normalization over whatever real geometry we have (filter bad values/outliers)
            if coords_real or poly_points:
                xs_all = [p[0] for p in coords_real.values()] + [p[0] for p in poly_points]
                ys_all = [p[1] for p in coords_real.values()] + [p[1] for p in poly_points]
                minxr, maxxr = min(xs_all), max(xs_all)
                minyr, maxyr = min(ys_all), max(ys_all)
                spanxr = max(1.0, maxxr - minxr)
                spanyr = max(1.0, maxyr - minyr)
                # Drop extreme outliers beyond ~5 sigma from median (basic safeguard)
                try:
                    import statistics as _st
                    mx = _st.median(xs_all); my = _st.median(ys_all)
                    sx = _st.pstdev(xs_all) or 1.0; sy = _st.pstdev(ys_all) or 1.0
                    limx = 5.0*sx; limy = 5.0*sy
                    def _okx(x: float) -> bool: return abs(x - mx) <= limx
                    def _oky(y: float) -> bool: return abs(y - my) <= limy
                    coords_real = {n:(x,y) for n,(x,y) in coords_real.items() if _okx(x) and _oky(y)}
                    poly_points = [(x,y) for (x,y) in poly_points if _okx(x) and _oky(y)]
                    if coords_real or poly_points:
                        xs_all = [p[0] for p in coords_real.values()] + [p[0] for p in poly_points]
                        ys_all = [p[1] for p in coords_real.values()] + [p[1] for p in poly_points]
                        minxr, maxxr = min(xs_all), max(xs_all)
                        minyr, maxyr = min(ys_all), max(ys_all)
                        spanxr = max(1.0, maxxr - minxr)
                        spanyr = max(1.0, maxyr - minyr)
                except Exception:
                    pass
                def _norm_geom_to_unit(x: float, y: float) -> tuple[float, float]:
                    return (x - minxr) / spanxr, (y - minyr) / spanyr
                coords_real_unit = {n: _norm_geom_to_unit(x, y) for n, (x, y) in coords_real.items()}
            else:
                coords_real_unit = {}

            # Merge: BFS unit, overridden by any real-unit positions
            coords_unit: dict[str, tuple[float, float]] = dict(coords_bfs)
            for n, xy in coords_real_unit.items():
                coords_unit[n] = xy
            # Pull nodes with no real coords toward average of real neighbors (jittered)
            # Soft relaxation only on synthetic nodes, never move anchored
            try:
                for _ in range(8):
                    for n in nodes:
                        if n in coords_real_unit:
                            continue  # locked anchors
                        neigh_u = [coords_unit.get(v) for v in adj.get(n, set())]
                        neigh_u = [(x, y) for (x, y) in neigh_u if x is not None and y is not None]
                        if not neigh_u:
                            continue
                        ax = sum(x for x, _ in neigh_u) / float(len(neigh_u))
                        ay = sum(y for _, y in neigh_u) / float(len(neigh_u))
                        px, py = coords_unit.get(n, (ax, ay))
                        # small step toward neighbors average
                        s = 0.2
                        nx = px + s * (ax - px)
                        ny = py + s * (ay - py)
                        coords_unit[n] = (min(1.0, max(0.0, nx)), min(1.0, max(0.0, ny)))
            except Exception:
                pass

            # Normalize polylines to unit using geometry extents if available
            polylines_unit: dict[str, list[list[tuple[float, float]]]] = {}
            try:
                if coords_real or poly_points:
                    def _norm_geom_to_unit(x: float, y: float) -> tuple[float, float]:
                        return (x - minxr) / spanxr, (y - minyr) / spanyr  # type: ignore[name-defined]
                    for k, plist in section_polylines.items():
                        out_list: list[list[tuple[float, float]]] = []
                        for pts in plist:
                            out_list.append([_norm_geom_to_unit(px, py) for (px, py) in pts])
                        if out_list:
                            polylines_unit[k] = out_list
            except Exception:
                pass

            # If we have tertiary nodes and a base-edge polyline, split the polyline at
            # the junction nearest to the tertiary anchor so each branch gets its shape.
            try:
                for k, tert_set in tert_by_edge.items():
                    if k not in polylines_unit:
                        continue
                    try:
                        u, v = k.split('|', 1)
                    except Exception:
                        continue
                    for tert in list(tert_set):
                        if tert not in coords_unit:
                            continue
                        tx, ty = coords_unit[tert]
                        base_list = polylines_unit.get(k) or []
                        add_uv: list[list[tuple[float, float]]] = []
                        add_tv: list[list[tuple[float, float]]] = []
                        tol = 0.02  # unit-space tolerance for tee snapping
                        tol2 = tol * tol
                        for pts in base_list:
                            if not pts:
                                continue
                            # find nearest poly point to tertiary
                            best_i = 0; best_d = 1e9
                            for i, (px, py) in enumerate(pts):
                                d = (px - tx) * (px - tx) + (py - ty) * (py - ty)
                                if d < best_d:
                                    best_d = d; best_i = i
                            if best_d <= tol2:
                                left = pts[:best_i+1]
                                right = pts[best_i:]
                                if left and right:
                                    add_uv.append(left)
                                    add_tv.append(right)
                        # attach splits and leave original polyline as-is (renderer prefers edge-specific ones)
                        if add_uv:
                            polylines_unit.setdefault(edge_key(u, tert), []).extend(add_uv)
                        if add_tv:
                            polylines_unit.setdefault(edge_key(tert, v), []).extend(add_tv)
            except Exception:
                pass

            # If nodes lack real coords, try to infer their positions from polyline endpoints
            if polylines_unit:
                inferred: dict[str, list[tuple[float, float]]] = {}
                for (u, v) in edges:
                    k = edge_key(u, v)
                    plist = polylines_unit.get(k)
                    if not plist:
                        continue
                    # Choose ends per edge to match BFS orientation
                    bu = coords_bfs.get(u, (0.0, 0.0))
                    bv = coords_bfs.get(v, (1.0, 1.0))
                    for pts in plist:
                        if not pts:
                            continue
                        a = pts[0]; b = pts[-1]
                        try:
                            s1 = (bu[0]-a[0])**2 + (bu[1]-a[1])**2 + (bv[0]-b[0])**2 + (bv[1]-b[1])**2
                            s2 = (bu[0]-b[0])**2 + (bu[1]-b[1])**2 + (bv[0]-a[0])**2 + (bv[1]-a[1])**2
                            if s1 <= s2:
                                inferred.setdefault(u, []).append(a)
                                inferred.setdefault(v, []).append(b)
                            else:
                                inferred.setdefault(u, []).append(b)
                                inferred.setdefault(v, []).append(a)
                        except Exception:
                            continue
                # Average candidates
                for n, lst in inferred.items():
                    if n in coords_real_unit:
                        continue
                    if not lst:
                        continue
                    ax = sum(x for x, _ in lst) / float(len(lst))
                    ay = sum(y for _, y in lst) / float(len(lst))
                    coords_unit[n] = (ax, ay)

            # Prepare result
            out["coords_unit"] = coords_unit
            out["edges"] = edges
            out["polylines_unit"] = polylines_unit
            # Heuristic Y flip: diagram-like (no geo) tends to be y-down. If we have
            # polylines, sample gradients; otherwise fall back to diagram default.
            y_down = True
            try:
                if used_geo:
                    y_down = False
                else:
                    grads = []
                    for plist in polylines_unit.values():
                        for pts in plist:
                            for i in range(len(pts)-1):
                                (x0, y0), (x1, y1) = pts[i], pts[i+1]
                                if abs(x1 - x0) + abs(y1 - y0) > 1e-6:
                                    grads.append(y1 - y0)
                    if grads:
                        # If most segments increase y downward (positive), treat as y-down
                        pos = sum(1 for g in grads if g > 0)
                        neg = sum(1 for g in grads if g < 0)
                        y_down = pos >= neg
            except Exception:
                pass
            out["y_down"] = y_down
            # Devices
            out["bus_sources"] = [b for b in bus_sources if b in nodes_set]
            out["bus_loads"] = bus_loads
            out["bus_shunts"] = bus_shunts
            out["inline_devs"] = inline_devs
            # Diagnostics
            try:
                anchors_count = len({n for n in nodes if n in coords_real_unit})
                poly_sec_count = sum(1 for k,v in (polylines_unit or {}).items() if v)
                poly_pts_count = sum(len(pts) for plist in (polylines_unit or {}).values() for pts in plist)
                synth_nodes = [n for n in nodes if n not in coords_real_unit]
                out["diag"] = {
                    "anchors": anchors_count,
                    "poly_sections": poly_sec_count,
                    "poly_points": poly_pts_count,
                    "synthetic_nodes": len(synth_nodes),
                    "y_down": y_down,
                }
            except Exception:
                pass
            return out
        except Exception as e:
            out["error"] = str(e)
            return out

    def _render_island_map_from_data(self, data: dict) -> None:
        # Fast path: draw from precomputed unit coordinates/polylines
        try:
            coords_unit = data.get('coords_unit')
            edges = data.get('edges')
            polylines_unit = data.get('polylines_unit') or {}
            isl = data.get('isl')
        except Exception:
            coords_unit = None; edges = None; polylines_unit = {}; isl = None

        if isinstance(coords_unit, dict) and isinstance(edges, list):
            try:
                if not hasattr(self, 'island_map') or self.island_map is None:
                    return
                cv = self.island_map
                cv.delete('all')
                w = max(100, int(cv.winfo_width() or 0))
                h = max(100, int(cv.winfo_height() or 0))
                # If we have no geometry (anchors and polylines absent), draw manual orthogonal layout
                try:
                    diag = data.get('diag') if isinstance(data, dict) else None
                    if isinstance(diag, dict) and diag.get('anchors', 0) == 0 and diag.get('poly_sections', 0) == 0:
                        self._draw_island_map_orthogonal(data, w, h)
                        return
                except Exception:
                    pass
                if not coords_unit:
                    # Nothing to draw — provide a gentle placeholder
                    try:
                        cv.create_text(w//2, h//2, text='(no buses in this island)', fill=self.COL.get('TEXT', '#334155'))
                    except Exception:
                        pass
                    return

                # Build bounds in unit space from nodes + polylines
                xs = [xy[0] for xy in coords_unit.values()] if coords_unit else [0.0, 1.0]
                ys = [xy[1] for xy in coords_unit.values()] if coords_unit else [0.0, 1.0]
                try:
                    for plist in polylines_unit.values():
                        for pts in plist:
                            for (ux, uy) in pts:
                                xs.append(ux); ys.append(uy)
                except Exception:
                    pass
                minx, maxx = (min(xs), max(xs)) if xs else (0.0, 1.0)
                miny, maxy = (min(ys), max(ys)) if ys else (0.0, 1.0)
                spanx = max(1e-9, maxx - minx)
                spany = max(1e-9, maxy - miny)
                pad = 20
                y_down = bool(data.get('y_down', True))
                # Preserve aspect ratio by using a single scale (letterbox)
                sx = (w - 2*pad) / spanx
                sy = (h - 2*pad) / spany
                s = min(sx, sy)
                ox = (w - s*spanx) * 0.5
                oy = (h - s*spany) * 0.5
                def _to_canvas(x: float, y: float) -> tuple[float, float]:
                    nx = ox + s * (x - minx)
                    if y_down:
                        ny = oy + s * (y - miny)
                    else:
                        ny = oy + s * (maxy - y)
                    return nx, ny

                coords_px: dict[str, tuple[float, float]] = {}
                for n, (ux, uy) in coords_unit.items():
                    coords_px[n] = _to_canvas(ux, uy)

                # Precompute canvas-space polylines
                polylines_px: dict[str, list[list[tuple[float, float]]]] = {}
                try:
                    for k, plist in polylines_unit.items():
                        out_list: list[list[tuple[float, float]]] = []
                        for pts in plist:
                            out_list.append([_to_canvas(px, py) for (px, py) in pts])
                        if out_list:
                            polylines_px[k] = out_list
                except Exception:
                    pass

                # Snap polyline endpoints to node anchors with tolerance in pixel space
                try:
                    tol = max(w, h) * 0.004  # ~0.4% of canvas
                    tol2 = tol * tol
                    max_snap_from = 0.0; max_snap_to = 0.0
                    for k, plist in list(polylines_px.items()):
                        try:
                            u, v = k.split('|', 1)
                        except Exception:
                            continue
                        au = coords_px.get(u); av = coords_px.get(v)
                        if not plist or (au is None and av is None):
                            continue
                        for pts in plist:
                            if not pts:
                                continue
                            # snap first to u, last to v when close
                            if au is not None:
                                dx = pts[0][0] - au[0]; dy = pts[0][1] - au[1]
                                if dx*dx + dy*dy <= tol2:
                                    d = (dx*dx + dy*dy) ** 0.5
                                    if d > max_snap_from: max_snap_from = d
                                    pts[0] = au
                            if av is not None:
                                dx = pts[-1][0] - av[0]; dy = pts[-1][1] - av[1]
                                if dx*dx + dy*dy <= tol2:
                                    d = (dx*dx + dy*dy) ** 0.5
                                    if d > max_snap_to: max_snap_to = d
                                    pts[-1] = av
                except Exception:
                    pass

                # If no anchors and no polylines, overlay a fallback notice
                try:
                    diag = data.get('diag') if isinstance(data, dict) else None
                    if isinstance(diag, dict) and (diag.get('anchors', 0) == 0) and (diag.get('poly_sections', 0) == 0):
                        cv.create_text(10, 10, text='No diagram geometry — fallback layout', anchor='nw', fill=self.COL.get('TEXT', '#334155'))
                    # Log snap distances summary
                    if isinstance(diag, dict):
                        self._append_log(f"Snap dists — from≈{max_snap_from:.2f}px, to≈{max_snap_to:.2f}px")
                        self._append_log(f"Canvas scale s={s:.3f}, letterbox=({ox:.1f},{oy:.1f}), bbox=[{minx:.3f}..{maxx:.3f}]x[{miny:.3f}..{maxy:.3f}] flip_y={y_down}")
                except Exception:
                    pass

                edge_color = self.COL.get('MUTED', '#94A3B8')
                _draw_count = 0
                for (u, v) in edges:
                    if u in coords_px and v in coords_px:
                        x1, y1 = coords_px[u]; x2, y2 = coords_px[v]
                        k = "|".join(sorted((u, v)))
                        if k in polylines_px:
                            for pts in polylines_px[k]:
                                # Orient roughly towards u->v
                                first_end = None
                                try:
                                    if pts:
                                        d_u = (pts[0][0]-x1)**2 + (pts[0][1]-y1)**2
                                        d_v = (pts[0][0]-x2)**2 + (pts[0][1]-y2)**2
                                        first_end = 'u' if d_u <= d_v else 'v'
                                except Exception:
                                    pass
                                seq: list[tuple[float, float]] = [coords_px[u]] + (list(reversed(pts)) if first_end == 'v' else pts) + [coords_px[v]]
                                flat: list[float] = []
                                for (px, py) in seq:
                                    flat.extend([px, py])
                                cv.create_line(*flat, fill=edge_color, width=1)
                                _draw_count += 1
                                if _draw_count % 200 == 0:
                                    try:
                                        cv.update()
                                    except Exception:
                                        pass
                        else:
                            cv.create_line(x1, y1, x2, y2, fill=edge_color, width=1)
                            _draw_count += 1
                            if _draw_count % 200 == 0:
                                try:
                                    cv.update()
                                except Exception:
                                    pass

                # Devices (precomputed in data) ---------------------------------
                # Build helper to get a representative path per edge
                def _path_for_edge(u: str, v: str) -> list[tuple[float, float]]:
                    k = "|".join(sorted((u, v)))
                    pts_variants = polylines_px.get(k)
                    if pts_variants:
                        pts = pts_variants[0]
                        try:
                            ux, uy = coords_px[u]; vx, vy = coords_px[v]
                            if pts:
                                d_u = (pts[0][0]-ux)**2 + (pts[0][1]-uy)**2
                                d_v = (pts[0][0]-vx)**2 + (pts[0][1]-vy)**2
                                if d_u > d_v:
                                    pts = list(reversed(pts))
                        except Exception:
                            pass
                        return [coords_px[u]] + list(pts) + [coords_px[v]]
                    return [coords_px[u], coords_px[v]]

                def _line_point_at_fraction(pts: list[tuple[float, float]], t: float) -> tuple[float, float, float, float, float, float]:
                    if t < 0.0: t = 0.0
                    if t > 1.0: t = 1.0
                    import math
                    total = 0.0
                    seglen: list[float] = []
                    for i in range(len(pts)-1):
                        x0,y0 = pts[i]; x1_,y1_ = pts[i+1]
                        L = math.hypot(x1_-x0, y1_-y0)
                        seglen.append(L); total += L
                    if total <= 1e-9:
                        x,y = pts[0]
                        return x,y,1.0,0.0,0.0,-1.0
                    target = t * total
                    acc = 0.0
                    for i in range(len(pts)-1):
                        L = seglen[i]
                        x0,y0 = pts[i]; x1_,y1_ = pts[i+1]
                        if acc + L >= target or i == len(pts)-2:
                            u = 0.0 if L <= 1e-9 else (target - acc) / L
                            x = x0 + u * (x1_-x0)
                            y = y0 + u * (y1_-y0)
                            tx = 1.0 if L <= 1e-9 else (x1_-x0)/L
                            ty = 0.0 if L <= 1e-9 else (y1_-y0)/L
                            nx = -ty; ny = tx
                            return x,y,tx,ty,nx,ny
                        acc += L
                    x,y = pts[-1]
                    return x,y,1.0,0.0,0.0,-1.0

                # Device data
                bus_sources = set(data.get('bus_sources') or [])
                bus_loads = dict(data.get('bus_loads') or {})
                bus_shunts = dict(data.get('bus_shunts') or {})
                inline_devs = dict(data.get('inline_devs') or {})

                # Glyphs using theme colors
                def draw_switch_at(x: float, y: float, nx: float, ny: float, size: float = 5.0):
                    px = nx * size; py = ny * size
                    cv.create_line(x - px, y - py, x + px, y + py, fill=self.COL.get('TEXT', '#334155'), width=2)
                def draw_xfmr_at(x: float, y: float, nx: float, ny: float, r: float = 3.0, gap: float = 3.0):
                    px = nx * gap; py = ny * gap
                    for sx, sy in ((x - px, y - py), (x + px, y + py)):
                        cv.create_oval(sx - r, sy - r, sx + r, sy + r, outline=self.COL.get('ACCENT', '#7C3AED'), fill=self.COL.get('ACCENT_SOFT', '#EDE9FE'), width=1.5)
                def draw_load_near_bus(x: float, y: float, idx: int):
                    off = 10 + idx * 8; s = 4
                    cv.create_rectangle(x + off - s, y - s, x + off + s, y + s, outline=self.COL.get('TEXT', '#334155'), fill=self.COL.get('ACCENT_SOFT', '#EDE9FE'))
                def draw_shunt_near_bus(x: float, y: float, idx: int):
                    base = 10 + idx * 10; top = y - base - 4; bot = y - base + 4; cxm = x - 4; cxp = x + 4
                    cv.create_line(cxm, top, cxm, bot, fill=self.COL.get('TEXT', '#334155'), width=2)
                    cv.create_line(cxp, top, cxp, bot, fill=self.COL.get('TEXT', '#334155'), width=2)
                    cv.create_line(cxm, bot, cxp, bot, fill=self.COL.get('TEXT', '#334155'), width=1)
                def draw_source_near_bus(x: float, y: float):
                    x = x - 12; r = 6
                    cv.create_oval(x - r, y - r, x + r, y + r, outline=self.COL.get('DANGER', '#EF4444'), width=2)
                    bolt = [(x-1, y-3), (x+1, y-3), (x-1, y+1), (x+1, y+1)]
                    cv.create_polygon(bolt, fill=self.COL.get('DANGER', '#EF4444'))

                # Inline devices along edge paths
                for k, devs_for_edge in inline_devs.items():
                    try:
                        u, v = k.split('|', 1)
                    except Exception:
                        continue
                    if u not in coords_px or v not in coords_px:
                        continue
                    path_pts = _path_for_edge(u, v)
                    loc_to_t = {'from': 0.2, 'middle': 0.5, 'to': 0.8}
                    stack = 0
                    for d in devs_for_edge:
                        loc = (d.get('loc') or 'middle').lower()
                        t = loc_to_t.get(loc, 0.5)
                        x, y, tx, ty, nx, ny = _line_point_at_fraction(path_pts, t)
                        off = (stack % 3 - 1) * 6.0
                        sx = x + nx * off; sy = y + ny * off
                        if d.get('type') == 'switch':
                            draw_switch_at(sx, sy, nx, ny)
                        elif d.get('type') == 'xfmr':
                            draw_xfmr_at(sx, sy, nx, ny)
                        stack += 1

                # Bus-attached devices
                for b, cnt in bus_loads.items():
                    if b in coords_px:
                        x, y = coords_px[b]
                        for i in range(cnt):
                            draw_load_near_bus(x, y, i)
                for b, cnt in bus_shunts.items():
                    if b in coords_px:
                        x, y = coords_px[b]
                        for i in range(cnt):
                            draw_shunt_near_bus(x, y, i)
                for b in sorted(bus_sources):
                    if b in coords_px:
                        x, y = coords_px[b]
                        draw_source_near_bus(x, y)

                # Draw nodes
                node_fill = self.COL.get('CARD', '#FFFFFF')
                node_stroke = self.COL.get('TEXT', '#334155')
                try:
                    ctx = get_island_context() or {}
                    slack_per_island: dict[int, str] = dict(ctx.get('slack_per_island', {}))
                    slack = slack_per_island.get(isl)
                except Exception:
                    slack = None
                for n, (x, y) in coords_px.items():
                    r = 3 if n != slack else 5
                    cv.create_oval(x - r, y - r, x + r, y + r, fill=node_fill, outline=node_stroke, width=1.5)

                # Scrollregion + center
                try:
                    bbox = cv.bbox('all')
                    if bbox:
                        x1, y1, x2, y2 = bbox
                        margin = 20
                        cv.configure(scrollregion=(x1 - margin, y1 - margin, x2 + margin, y2 + margin))
                        cx = (x1 + x2) / 2.0; cy = (y1 + y2) / 2.0
                        dx = (w / 2.0) - cx; dy = (h / 2.0) - cy
                        cv.move('all', dx, dy)
                        self._map_fitted = True
                except Exception:
                    pass
                return
            except Exception:
                # fall through to legacy path
                pass

        # Fallback to legacy synchronous drawer if precomputed data missing
        try:
            isl = data.get("isl")
            bus_to_island = data.get("bus_to_island") or {}
            if not isinstance(bus_to_island, dict):
                bus_to_island = {}
            self._draw_island_map(isl, bus_to_island)
        except Exception:
            try:
                if hasattr(self, 'island_map') and self.island_map is not None:
                    self.island_map.delete('all')
            except Exception:
                pass

    def _draw_island_map(self, isl: int | None, bus_to_island: dict[str, int]):
        if not hasattr(self, 'island_map') or self.island_map is None:
            return
        cv = self.island_map
        cv.delete('all')
        w = max(100, int(cv.winfo_width() or 0))
        h = max(100, int(cv.winfo_height() or 0))

        # Gather island nodes
        if isl is None:
            # Nothing selected
            return
        nodes = sorted([b for b, i in bus_to_island.items() if i == isl])
        if not nodes:
            return

        # Build adjacency from the input file, then subgraph to selected island
        try:
            from Modules.IslandChecker import _read_xml as _is_read_xml  # type: ignore
            from Modules.IslandChecker import _build_graph as _is_build_graph  # type: ignore
            in_path = Path(self.in_path.get() or "").expanduser()
            root = _is_read_xml(in_path)
            adj_full, _, _ = _is_build_graph(root)
        except Exception:
            adj_full = {n: set() for n in nodes}

        adj = {n: set(v for v in adj_full.get(n, set()) if v in nodes) for n in nodes}

        # Try to collect real coordinates (best-effort, optional)
        coords_real: dict[str, tuple[float, float]] = {}
        # Optional polylines for sections with bends (IntermediatePoints)
        # Map undirected edge key -> list of {from, to, pts[(x,y), ...]}
        section_polylines: dict[frozenset[str], list[dict[str, Any]]] = {}
        try:
            import xml.etree.ElementTree as ET
            from Modules.General import safe_name as _safe
            def _num(x: str | None) -> float | None:
                try:
                    return float(x) if x not in (None, "") else None
                except Exception:
                    return None
            def _is_zero_xy(x: float | None, y: float | None) -> bool:
                try:
                    return bool(x is not None and y is not None and abs(x) == 0.0 and abs(y) == 0.0)
                except Exception:
                    return False
            # Gather coordinates from two sources:
            #  1) Any element that has NodeID + X/Y (node-centric coords)
            #  2) Sections with From*/To* X/Y (endpoint coords)
            #  3) Sections with IntermediatePoints for bends (polyline)
            in_path = Path(self.in_path.get() or "").expanduser()
            root = ET.fromstring(in_path.read_text(encoding='utf-8', errors='ignore'))
            # 1) Node-centric coords
            for elem in root.iter():
                nid_raw = (elem.findtext('NodeID') or '').strip()
                if not nid_raw:
                    continue
                nx = _num(elem.findtext('X') or elem.findtext('Longitude'))
                ny = _num(elem.findtext('Y') or elem.findtext('Latitude'))
                # Skip (0,0) to avoid visual blow-ups
                if nx is not None and ny is not None and not _is_zero_xy(nx, ny):
                    nid = _safe(nid_raw)
                    if nid:
                        coords_real.setdefault(nid, (nx, ny))
            # 2) Section endpoint coords
            for sec in root.findall('.//Sections/Section'):
                fb_raw = (sec.findtext('FromNodeID') or '').strip(); tb_raw = (sec.findtext('ToNodeID') or '').strip()
                fb = _safe(fb_raw) if fb_raw else ''
                tb = _safe(tb_raw) if tb_raw else ''
                # Collect all tag/text pairs under the section
                kv: dict[str, str] = {}
                for elem in sec.iter():
                    tag = elem.tag.split('}')[-1].lower()
                    txt = (elem.text or '').strip()
                    if txt:
                        kv[tag] = txt
                    for ak, av in elem.attrib.items():
                        k2 = f"{tag}.{ak}".lower()
                        if av and k2 not in kv:
                            kv[k2] = av
                def pick(prefixes: list[str]) -> tuple[float | None, float | None]:
                    x_val = None; y_val = None
                    keys = list(kv.keys())
                    # Prefer explicit longitude/latitude
                    for k in keys:
                        lk = k.replace('_','').replace('-','')
                        if any(px in lk for px in prefixes):
                            if any(ax in lk for ax in ['long','lon']):
                                xv = _num(kv[k]);
                                if xv is not None: x_val = xv
                            if 'lat' in lk:
                                yv = _num(kv[k]);
                                if yv is not None: y_val = yv
                    # Generic X/Y for the given endpoint
                    if x_val is None:
                        for k in keys:
                            lk = k.replace('_','').replace('-','')
                            if any(px in lk for px in prefixes) and any(ax in lk for ax in ['x','xcoord','xcoordinate','posx','positionx','xpos']):
                                xv = _num(kv[k]);
                                if xv is not None: x_val = xv; break
                    if y_val is None:
                        for k in keys:
                            lk = k.replace('_','').replace('-','')
                            if any(px in lk for px in prefixes) and any(ay in lk for ay in ['y','ycoord','ycoordinate','posy','positiony','ypos']):
                                yv = _num(kv[k]);
                                if yv is not None: y_val = yv; break
                    # Last resort: plain X/Y at section scope
                    if x_val is None:
                        x_val = _num(kv.get('x'))
                    if y_val is None:
                        y_val = _num(kv.get('y'))
                    return x_val, y_val
                fx, fy = pick(['from'])
                tx, ty = pick(['to'])
                if fb and fb in nodes and fx is not None and fy is not None and not _is_zero_xy(fx, fy):
                    coords_real.setdefault(fb, (fx, fy))
                if tb and tb in nodes and tx is not None and ty is not None and not _is_zero_xy(tx, ty):
                    coords_real.setdefault(tb, (tx, ty))

                # 3) Polyline (IntermediatePoints) — record bends only, never draw nodes on bends
                try:
                    ip_elem = sec.find('IntermediatePoints')
                    if ip_elem is not None and fb and tb:
                        pts: list[tuple[float, float]] = []
                        for p in ip_elem.findall('Point'):
                            px = _num((p.findtext('X') or '').strip())
                            py = _num((p.findtext('Y') or '').strip())
                            if px is None or py is None:
                                continue
                            if _is_zero_xy(px, py):
                                continue
                            pts.append((px, py))
                        if pts:
                            key = frozenset({fb, tb})
                            section_polylines.setdefault(key, []).append({'from': fb, 'to': tb, 'pts': pts})
                except Exception:
                    pass
        except Exception:
            pass

        # Allow UI to process events so spinner animates between heavy phases
        try:
            cv.update_idletasks()
        except Exception:
            pass

        # Always generate a fallback layered one-line diagram (columns by BFS depth)
        ctx = get_island_context() or {}
        slack_per_island: dict[int, str] = dict(ctx.get('slack_per_island', {}))
        root_node = slack_per_island.get(isl) or (nodes[0] if nodes else None)
        from collections import deque as _dq
        depth: dict[str, int] = {}
        if root_node:
            q = _dq([root_node]); depth[root_node] = 0
            while q:
                u = q.popleft()
                for v in sorted(adj.get(u, set())):
                    if v not in depth:
                        depth[v] = depth[u] + 1
                        q.append(v)
        for n in nodes:
            depth.setdefault(n, 0)
        columns: dict[int, list[str]] = {}
        for n in nodes:
            columns.setdefault(depth[n], []).append(n)
        for col in columns.values():
            col.sort()
        coords_bfs: dict[str, tuple[float, float]] = {}
        for dlevel, col_nodes in columns.items():
            k = max(1, len(col_nodes))
            for i, n in enumerate(col_nodes):
                x = float(dlevel)
                y = float(i) / float(k - 1 if k > 1 else 1)
                coords_bfs[n] = (x, y)
        # Normalize BFS x-depth to unit [0..1] so it blends well with real XY that
        # we also scale to unit space; prevents UTM-vs-depth scale clashes.
        try:
            dmax = max((xy[0] for xy in coords_bfs.values()), default=1.0)
            if dmax <= 0:
                dmax = 1.0
            for n, (x, y) in list(coords_bfs.items()):
                coords_bfs[n] = (x / dmax, y)
        except Exception:
            pass

        # Merge with scale separation: real-world coords are first normalized to unit space
        # using only the real-coord extent, so BFS (0..1) doesn't get crushed by UTM-scale values.
        if coords_real:
            xsr = [p[0] for p in coords_real.values()]
            ysr = [p[1] for p in coords_real.values()]
            minxr, maxxr = min(xsr), max(xsr)
            minyr, maxyr = min(ysr), max(ysr)
            spanxr = max(1.0, maxxr - minxr)
            spanyr = max(1.0, maxyr - minyr)
            def _norm_real_to_unit(x: float, y: float) -> tuple[float, float]:
                return (x - minxr) / spanxr, (y - minyr) / spanyr
            coords_real_unit = {n: _norm_real_to_unit(x, y) for n, (x, y) in coords_real.items()}
        else:
            coords_real_unit = {}

        # Unit-space coords for all nodes: start with BFS, override with any real
        coords_unit: dict[str, tuple[float, float]] = dict(coords_bfs)
        for n, xy in coords_real_unit.items():
            coords_unit[n] = xy
        # If a node lacks real coords but is adjacent to nodes that have them,
        # pull it toward the average of its real-neighbor positions with tiny jitter.
        try:
            import math as _m
            for n in nodes:
                if n in coords_real_unit:
                    continue
                neigh_u = [coords_real_unit[v] for v in adj.get(n, set()) if v in coords_real_unit]
                if not neigh_u:
                    continue
                ax = sum(x for x, _ in neigh_u) / float(len(neigh_u))
                ay = sum(y for _, y in neigh_u) / float(len(neigh_u))
                # stable tiny jitter from hash to avoid overplotting
                h = hash(n)
                jx = ((h & 0x3F) / 63.0 - 0.5) * 0.02  # +/-0.01
                jy = (((h >> 6) & 0x3F) / 63.0 - 0.5) * 0.02
                coords_unit[n] = (ax + jx, ay + jy)
        except Exception:
            pass

        # Allow UI to process events before mapping to canvas
        try:
            cv.update_idletasks()
        except Exception:
            pass

        # Now normalize unit-space coords to canvas space; include polylines with real coords also normalized to unit first
        xs = [p[0] for p in coords_unit.values()]
        ys = [p[1] for p in coords_unit.values()]
        # Include normalized polylines' points in bounds (even if endpoints lack node coords)
        polylines_unit: dict[frozenset[str], list[list[tuple[float, float]]]] = {}
        try:
            for key, plist in section_polylines.items():
                out_list: list[list[tuple[float, float]]] = []
                for poly in plist:
                    pts_raw = poly.get('pts', [])
                    pts_u = [(_norm_real_to_unit(px, py) if coords_real else (px, py)) for (px, py) in pts_raw]
                    out_list.append(pts_u)
                    for (ux, uy) in pts_u:
                        xs.append(ux); ys.append(uy)
                if out_list:
                    polylines_unit[key] = out_list
        except Exception:
            pass

        minx, maxx = (min(xs), max(xs)) if xs else (0.0, 1.0)
        miny, maxy = (min(ys), max(ys)) if ys else (0.0, 1.0)
        spanx = max(1e-6, maxx - minx)
        spany = max(1e-6, maxy - miny)
        pad = 20
        def _to_canvas(x: float, y: float) -> tuple[float, float]:
            nx = pad + (x - minx) / spanx * (w - 2*pad)
            ny = pad + (y - miny) / spany * (h - 2*pad)
            return nx, ny
        coords: dict[str, tuple[float, float]] = {}
        for n, (x, y) in coords_unit.items():
            coords[n] = _to_canvas(x, y)

        # Precompute canvas-space polylines for edges with bends
        polylines_norm: dict[frozenset[str], list[list[tuple[float, float]]]] = {}
        try:
            for key, plist in polylines_unit.items():
                out_list: list[list[tuple[float, float]]] = []
                for pts_u in plist:
                    out_list.append([_to_canvas(px, py) for (px, py) in pts_u])
                if out_list:
                    polylines_norm[key] = out_list
        except Exception:
            pass

        # Draw edges (tick UI periodically so spinner keeps animating)
        edge_color = self.COL.get('MUTED', '#94A3B8')
        _draw_count = 0
        for u in nodes:
            for v in adj.get(u, set()):
                if u < v and u in coords and v in coords:
                    key = frozenset({u, v})
                    x1, y1 = coords[u]; x2, y2 = coords[v]
                    # If we have a polyline for this edge, draw the bend(s)
                    if key in polylines_norm:
                        for pts in polylines_norm[key]:
                            # Determine orientation: the stored section may be fb->tb or tb->fb
                            # We only reverse the bend points if needed to keep u->v order consistent visually
                            first_end = None
                            try:
                                if pts:
                                    d_u = (pts[0][0]-x1)**2 + (pts[0][1]-y1)**2
                                    d_v = (pts[0][0]-x2)**2 + (pts[0][1]-y2)**2
                                    first_end = 'u' if d_u <= d_v else 'v'
                            except Exception:
                                pass
                            seq: list[tuple[float, float]] = []
                            if first_end == 'v':
                                seq = [coords[u]] + list(reversed(pts)) + [coords[v]]
                            else:
                                seq = [coords[u]] + pts + [coords[v]]
                            flat: list[float] = []
                            for (px, py) in seq:
                                flat.extend([px, py])
                            cv.create_line(*flat, fill=edge_color, width=1)
                            _draw_count += 1
                            if _draw_count % 200 == 0:
                                try:
                                    cv.update_idletasks()
                                except Exception:
                                    pass
                    else:
                        cv.create_line(x1, y1, x2, y2, fill=edge_color, width=1)
                        _draw_count += 1
                        if _draw_count % 200 == 0:
                            try:
                                cv.update_idletasks()
                            except Exception:
                                pass

        # ---- Device glyphs (CYME-like) -------------------------------------
        # We draw:
        #  - Voltage source(s): at source buses (from Substation Topo) inside this island
        #  - Shunts + Loads: attached to buses (placed next to the bus circle)
        #  - Switches + Transformers: inline on sections (at From/Middle/To along the path)
        #
        # Parsing is best-effort; if anything fails we skip gracefully.
        try:
            import xml.etree.ElementTree as ET
            from Modules.General import safe_name as _safe
            in_path = Path(self.in_path.get() or "").expanduser()
            xml_root = ET.fromstring(in_path.read_text(encoding='utf-8', errors='ignore'))

            nodes_set = set(nodes)

            # Find source buses from Substation Topo
            vs_nodes: set[str] = set()
            for topo in xml_root.findall('.//Topo'):
                ntype = (topo.findtext('NetworkType') or '').strip().lower()
                eq_mode = (topo.findtext('EquivalentMode') or '').strip()
                if ntype != 'substation' or eq_mode == '1':
                    continue
                srcs = topo.find('./Sources')
                if srcs is None:
                    continue
                for src in srcs.findall('./Source'):
                    nid = _safe(src.findtext('SourceNodeID') or '')
                    if nid:
                        vs_nodes.add(nid)

            # Bus-attached devices: count per bus for placement offsets
            bus_loads: dict[str, int] = {}
            bus_shunts: dict[str, int] = {}
            bus_sources: set[str] = set()

            # Inline devices per undirected edge key
            inline_devs: dict[frozenset[str], list[dict[str, Any]]] = {}

            # Helper to add inline device record
            def add_inline(dev_type: str, fb: str, tb: str, loc: str | None = None) -> None:
                if not fb or not tb:
                    return
                if fb not in nodes_set or tb not in nodes_set:
                    return
                key = frozenset({fb, tb})
                inline_devs.setdefault(key, []).append({
                    'type': dev_type,
                    'from': fb,
                    'to': tb,
                    'loc': (loc or 'middle').lower()
                })

            # Populate bus-attached and inline devices by scanning sections
            for sec in xml_root.findall('.//Sections/Section'):
                fb = _safe((sec.findtext('FromNodeID') or '').strip())
                tb = _safe((sec.findtext('ToNodeID') or '').strip())
                if not fb or not tb:
                    continue
                # Restrict to selected island endpoints
                if (fb not in nodes_set) and (tb not in nodes_set):
                    continue

                devs = sec.find('./Devices')
                if devs is None:
                    continue

                # Loads (Spot + Distributed) attach to From bus per sheet rules
                if devs.find('SpotLoad') is not None or devs.find('DistributedLoad') is not None:
                    if fb in nodes_set:
                        bus_loads[fb] = bus_loads.get(fb, 0) + 1

                # Shunt (Capacitor/Reactor) attach to From bus per sheet rules
                if devs.find('ShuntCapacitor') is not None or devs.find('ShuntReactor') is not None:
                    if fb in nodes_set:
                        bus_shunts[fb] = bus_shunts.get(fb, 0) + 1

                # Switch-like devices inline. Consider Location tag for approximate placement.
                for tag in ('Switch', 'Sectionalizer', 'Breaker', 'Fuse', 'Recloser'):
                    for d in devs.findall(tag):
                        loc = (d.findtext('Location') or 'Middle')
                        add_inline('switch', fb, tb, loc)
                # Some Miscellaneous codes behave as inline switches
                for d in devs.findall('Miscellaneous'):
                    did = ((d.findtext('DeviceID') or '').strip().upper())
                    if did in {'RB', 'LA'}:
                        loc = (d.findtext('Location') or 'Middle')
                        add_inline('switch', fb, tb, loc)

                # Transformers inline
                if devs.find('Transformer') is not None or devs.find('Regulator') is not None:
                    loc = 'Middle'
                    # if a Location field exists under Transformer, use it
                    xf = devs.find('Transformer')
                    if xf is not None:
                        loc = (xf.findtext('Location') or 'Middle')
                    add_inline('xfmr', fb, tb, loc)

            # Sources on buses for this island only
            for b in vs_nodes:
                if b in nodes_set:
                    bus_sources.add(b)

            # ---- Drawing helpers -------------------------------------------
            def line_point_at_fraction(pts: list[tuple[float, float]], t: float) -> tuple[float, float, float, float, float, float]:
                # clamp t
                if t < 0.0:
                    t = 0.0
                if t > 1.0:
                    t = 1.0
                # total length
                import math
                segs: list[tuple[float,float,float,float,float]] = []  # (x0,y0,x1,y1,len)
                total = 0.0
                for i in range(len(pts)-1):
                    x0,y0 = pts[i]; x1,y1 = pts[i+1]
                    dx = x1-x0; dy = y1-y0
                    L = math.hypot(dx, dy)
                    if L <= 1e-6:
                        continue
                    segs.append((x0,y0,x1,y1,L))
                    total += L
                if total <= 0.0:
                    # fallback: first point
                    x,y = pts[0]
                    return x,y, 1.0,0.0, 0.0,1.0
                target = t * total
                acc = 0.0
                for (x0,y0,x1,y1,L) in segs:
                    if acc + L >= target:
                        remain = target - acc
                        f = remain / L if L > 0 else 0.0
                        x = x0 + f * (x1 - x0)
                        y = y0 + f * (y1 - y0)
                        tx = (x1 - x0) / L
                        ty = (y1 - y0) / L
                        nx = -ty
                        ny = tx
                        return x,y, tx,ty, nx,ny
                    acc += L
                # if we fall through, return end of last
                x0,y0,x1,y1,L = segs[-1]
                tx = (x1 - x0) / L
                ty = (y1 - y0) / L
                return x1,y1, tx,ty, -ty,tx

            def draw_switch_at(x: float, y: float, nx: float, ny: float, size: float = 5.0):
                # small bar across the conductor (perpendicular)
                px = nx * size
                py = ny * size
                cv.create_line(x - px, y - py, x + px, y + py, fill=self.COL.get('TEXT', '#334155'), width=2)

            def draw_xfmr_at(x: float, y: float, nx: float, ny: float, r: float = 3.0, gap: float = 3.0):
                # two small circles offset along normal
                px = nx * gap; py = ny * gap
                for sx, sy in ((x - px, y - py), (x + px, y + py)):
                    cv.create_oval(sx - r, sy - r, sx + r, sy + r, outline=self.COL.get('ACCENT', '#7C3AED'), fill=self.COL.get('ACCENT_SOFT', '#EDE9FE'), width=1.5)

            def draw_load_near_bus(x: float, y: float, idx: int):
                # small square to the right, stacked
                off = 10 + idx * 8
                s = 4
                cv.create_rectangle(x + off - s, y - s, x + off + s, y + s, outline=self.COL.get('TEXT', '#334155'), fill=self.COL.get('ACCENT_SOFT', '#EDE9FE'))

            def draw_shunt_near_bus(x: float, y: float, idx: int):
                # capacitor symbol above the bus, stacked upward
                base = 10 + idx * 10
                top = y - base - 4
                bot = y - base + 4
                cxm = x - 4; cxp = x + 4
                cv.create_line(cxm, top, cxm, bot, fill=self.COL.get('TEXT', '#334155'), width=2)
                cv.create_line(cxp, top, cxp, bot, fill=self.COL.get('TEXT', '#334155'), width=2)
                cv.create_line(cxm, bot, cxp, bot, fill=self.COL.get('TEXT', '#334155'), width=1)

            def draw_source_near_bus(x: float, y: float):
                # offset to the left of bus circle to avoid overlap
                x = x - 12
                r = 6
                cv.create_oval(x - r, y - r, x + r, y + r, outline=self.COL.get('DANGER', '#EF4444'), width=2)
                # small lightning bolt inside
                bolt = [(x-1, y-3), (x+1, y-3), (x-1, y+1), (x+1, y+1)]
                cv.create_polygon(bolt, fill=self.COL.get('DANGER', '#EF4444'))

            # Draw inline devices using polylines when available; else straight path
            for key, dev_list in inline_devs.items():
                u, v = tuple(key)
                if u not in coords or v not in coords:
                    continue
                # Choose a representative path for this edge
                path_variants = polylines_norm.get(key)
                if path_variants and len(path_variants) > 0:
                    bends = path_variants[0]
                    path_pts = [coords[u]] + list(bends) + [coords[v]]
                else:
                    path_pts = [coords[u], coords[v]]

                # place devices; stack by slight offset along normal
                loc_to_t = {'from': 0.2, 'middle': 0.5, 'to': 0.8}
                stack = 0
                for d in dev_list:
                    t = loc_to_t.get((d.get('loc') or 'middle'), 0.5)
                    x, y, tx, ty, nx, ny = line_point_at_fraction(path_pts, t)
                    # small normal offset to avoid covering the line if multiple
                    off = (stack % 3 - 1) * 6.0  # -6,0,+6 repeating
                    sx = x + nx * off
                    sy = y + ny * off
                    if d.get('type') == 'switch':
                        draw_switch_at(sx, sy, nx, ny)
                    elif d.get('type') == 'xfmr':
                        draw_xfmr_at(sx, sy, nx, ny)
                    stack += 1

            # Draw bus devices
            for b, count in bus_loads.items():
                if b in coords:
                    x, y = coords[b]
                    for i in range(count):
                        draw_load_near_bus(x, y, i)
            for b, count in bus_shunts.items():
                if b in coords:
                    x, y = coords[b]
                    for i in range(count):
                        draw_shunt_near_bus(x, y, i)
            for b in sorted(bus_sources):
                if b in coords:
                    x, y = coords[b]
                    draw_source_near_bus(x, y)

        except Exception:
            # Silent failure for glyphs to keep map robust
            pass

        # Draw nodes
        node_fill = self.COL.get('ACCENT_SOFT', '#EDE9FE')
        node_stroke = self.COL.get('ACCENT', '#7C3AED')
        ctx = get_island_context() or {}
        slack_per_island: dict[int, str] = dict(ctx.get('slack_per_island', {}))
        slack = slack_per_island.get(isl)
        for n, (x, y) in coords.items():
            # Smaller circles; highlight slack slightly larger
            r = 3 if n != slack else 5
            cv.create_oval(x - r, y - r, x + r, y + r, fill=node_fill, outline=node_stroke, width=1.5)
            # Optional labels for debugging:
            # cv.create_text(x + r + 2, y, text=n, anchor='w', fill=self.COL.get('TEXT', '#334155'), font=(self.UI_FONT, max(8, self.UI_SIZE-3)))
        # Update scroll region and center content initially
        try:
            bbox = cv.bbox('all')
            if bbox:
                x1, y1, x2, y2 = bbox
                # Expand scroll region slightly for nice margins and recentre
                margin = 20
                cv.configure(scrollregion=(x1 - margin, y1 - margin, x2 + margin, y2 + margin))
                cx = (x1 + x2) / 2.0
                cy = (y1 + y2) / 2.0
                dx = (w / 2.0) - cx
                dy = (h / 2.0) - cy
                cv.move('all', dx, dy)
                self._map_fitted = True
        except Exception:
            pass

    def _on_map_wheel(self, event, delta: int | None = None):
        """Zoom the island map around the cursor position using mouse wheel."""
        try:
            cv = self.island_map
            if cv is None:
                return
            d = delta if delta is not None else getattr(event, 'delta', 0)
            if d == 0:
                return
            # Typical delta is +/-120 on Windows; scale factor per notch
            factor = 1.1 if d > 0 else (1/1.1)
            # Use canvas coordinates for stable zooming
            x = cv.canvasx(event.x)
            y = cv.canvasy(event.y)
            cv.scale('all', x, y, factor, factor)
            # Keep scrollregion in sync so panning works
            bbox = cv.bbox('all')
            if bbox:
                cv.configure(scrollregion=bbox)
        except Exception:
            pass

    # ----- Orthogonal fallback layout/draw (no geometry present) -------------
    def _draw_island_map_orthogonal(self, data: dict, w: int, h: int) -> None:
        cv = self.island_map
        if cv is None:
            return
        # Theme colors (single source)
        COL = self.COL
        edge_color     = COL.get('EDGE', 'black')
        text_color     = COL.get('TEXT', '#222')
        bus_fill       = COL.get('BUS', '#7ca5ff')
        source_fill    = COL.get('SOURCE', '#ff6961')
        xfr_fill       = COL.get('TRANSFORMER', '#8fd694')
        switch_fill    = COL.get('SWITCH', '#f7b267')
        load_fill      = COL.get('LOAD', '#b0b0b0')
        shunt_fill     = COL.get('SHUNT', '#7dcfb6')
        select_outline = COL.get('SELECT', '#6a0dad')
        bg_color       = COL.get('CANVAS_BG', '#ffffff')

        # Constants
        DX = 140; DY = 120; MARGIN = 40
        BUS_R = 6; SRC_R = 10
        SYMB_W, SYMB_H = 18, 12

        # Metadata and event bindings
        if not hasattr(self, 'meta'):
            self.meta = {}
        def tag_and_meta(item_id: int, typ: str, obj_id: str, name: str) -> None:
            try:
                cv.addtag_withtag('obj', item_id)
                cv.addtag_withtag(f'type:{typ}', item_id)
                cv.addtag_withtag(f'id:{obj_id}', item_id)
                self.meta[item_id] = {"type": typ, "id": obj_id, "name": name}
            except Exception:
                pass
        # Bind once
        try:
            if not getattr(self, '_obj_binds', False):
                cv.tag_bind('obj', '<Button-1>', self._on_obj_click)
                cv.tag_bind('obj', '<Enter>', self._on_obj_hover_enter)
                cv.tag_bind('obj', '<Leave>', self._on_obj_hover_leave)
                self._obj_binds = True
        except Exception:
            pass

        # Build node set and adjacency
        edges = list(data.get('edges') or [])
        nodes_set = set()
        for (u, v) in edges:
            nodes_set.add(u); nodes_set.add(v)
        bus_sources = set(data.get('bus_sources') or [])
        bus_loads = dict(data.get('bus_loads') or {})
        bus_shunts = dict(data.get('bus_shunts') or {})
        inline_devs = dict(data.get('inline_devs') or {})  # key "u|v" -> list[{type,loc}]

        adj: dict[str, set[str]] = {n: set() for n in nodes_set}
        for (u, v) in edges:
            adj.setdefault(u, set()).add(v)
            adj.setdefault(v, set()).add(u)

        # Choose trunk path
        def bfs_far(start: str) -> tuple[str, dict[str, str]]:
            from collections import deque
            q = deque([start]); seen = {start}; parent: dict[str, str] = {}
            last = start
            while q:
                u = q.popleft(); last = u
                for v in adj.get(u, set()):
                    if v not in seen:
                        seen.add(v); parent[v] = u; q.append(v)
            return last, parent

        if bus_sources:
            root = next(iter(bus_sources & nodes_set), next(iter(nodes_set)) if nodes_set else '')
        else:
            root = next(iter(nodes_set)) if nodes_set else ''
        a, _ = bfs_far(root)
        b, par = bfs_far(a)
        # Reconstruct path from a to b
        path = [b]
        while path[-1] in par:
            path.append(par[path[-1]])
        trunk = list(reversed(path))
        trunk = [n for n in trunk if n in nodes_set]
        trunk_set = set(trunk)
        # If trunk does not include root and root is a source, prepend root
        if bus_sources and root in nodes_set and root not in trunk_set:
            trunk = [root] + trunk
            trunk_set = set(trunk)

        # Assign grid positions
        gx: dict[str, int] = {}
        gy: dict[str, int] = {}
        for i, n in enumerate(trunk):
            gx[n] = i
            gy[n] = 0

        # Branch assignment per trunk node
        visited = set(trunk)
        side_toggle = -1  # alternate above/below
        for i, t in enumerate(trunk):
            # gather branch clusters from neighbors not on trunk
            neighs = [v for v in adj.get(t, set()) if v not in trunk_set]
            if not neighs:
                continue
            # For each immediate neighbor start a BFS to collect a branch cluster
            lane_idx = 1
            for start in neighs:
                if start in visited:
                    continue
                side_toggle *= -1
                side = side_toggle  # +1 above, -1 below (we'll flip later in mapping)
                # Alternate lanes ±lane_idx
                lane_offset = lane_idx
                lane_idx += 1
                col = gx[t] + lane_offset
                from collections import deque
                q = deque([(start, 1)])  # (node, depth)
                visited.add(start)
                gx[start] = col; gy[start] = side * 1
                while q:
                    u, d = q.popleft()
                    for v in adj.get(u, set()):
                        if v in visited or v in trunk_set:
                            continue
                        visited.add(v)
                        gx[v] = col; gy[v] = side * (d + 1)
                        q.append((v, d + 1))

        # Any remaining unplaced nodes (cycles etc.) — place near their neighbor's column
        for n in nodes_set:
            if n not in gx:
                # find any neighbor placed
                nn = next(iter([v for v in adj.get(n, set()) if v in gx]), None)
                if nn is None:
                    gx[n] = len(gx); gy[n] = 0
                else:
                    gx[n] = gx[nn]
                    gy[n] = gy[nn] + 1

        # Grid bounds
        min_gx = min(gx.values()) if gx else 0
        max_gx = max(gx.values()) if gx else 1
        min_gy = min(gy.values()) if gy else 0
        max_gy = max(gy.values()) if gy else 1
        gw = max(1, max_gx - min_gx)
        gh = max(1, max_gy - min_gy)

        # Aspect-preserving map from grid to canvas
        sx = (w - 2*MARGIN) / float(max(1, gw) * DX)
        sy = (h - 2*MARGIN) / float(max(1, gh) * DY)
        s = min(sx, sy)
        ox = (w - s * (gw * DX)) * 0.5
        oy = (h - s * (gh * DY)) * 0.5

        def to_px(pxg: int, pyg: int) -> tuple[float, float]:
            # Math Y-up: larger gy goes upward visually; invert for canvas (y-down)
            x = ox + s * ((pxg - min_gx) * DX)
            y = oy + s * ((max_gy - pyg) * DY)
            return x, y

        coords_px: dict[str, tuple[float, float]] = {n: to_px(gx[n], gy[n]) for n in nodes_set}

        # Route edges orthogonally (L-shapes)
        polylines_px: dict[str, list[tuple[float, float]]] = {}
        for (u, v) in edges:
            x1, y1 = coords_px[u]; x2, y2 = coords_px[v]
            if abs(x1 - x2) < 1e-6 or abs(y1 - y2) < 1e-6:
                poly = [(x1, y1), (x2, y2)]
            else:
                # vertical then horizontal to favor tee look from trunk
                mid = (x1, y2)
                poly = [(x1, y1), mid, (x2, y2)]
            polylines_px["|".join(sorted((u, v)))] = poly

        # Interactive tagging map
        self._map_meta = {}

        # Draw edges (black)
        for key, pts in polylines_px.items():
            flat = []
            for (x, y) in pts:
                flat.extend([x, y])
            item = cv.create_line(*flat, fill=edge_color, width=2)
            tag_and_meta(item, 'Edge', f'edge:{key}', f'Section {key.replace("|","-")}')

        # Inline devices along first segment from upstream (choose left-most as upstream)
        def place_on_first_segment(pts: list[tuple[float, float]]) -> tuple[float, float, bool]:
            if len(pts) < 2:
                return pts[0][0], pts[0][1], True
            (x0, y0), (x1, y1) = pts[0], pts[1]
            hx = (x0 + x1) * 0.5; hy = (y0 + y1) * 0.5
            horiz = abs(y1 - y0) < abs(x1 - x0)
            return hx, hy, horiz

        for k, devs in inline_devs.items():
            pts = polylines_px.get(k)
            if not pts:
                continue
            base_x, base_y, horiz = place_on_first_segment(pts)
            for idx, d in enumerate(devs):
                dtype = (d.get('type') or 'switch').lower()
                # slight offset to avoid overlap when multiple devices
                off = (idx % 3 - 1) * 8
                if horiz:
                    dx, dy = 0, off
                else:
                    dx, dy = off, 0
                x = base_x + dx; y = base_y + dy
                name = d.get('name') or (('Transformer' if dtype=='xfmr' else 'Switch') + f' {k}')
                if dtype == 'xfmr':
                    if horiz:
                        a = cv.create_oval(x-12, y-6, x-2, y+6, outline=text_color, width=2, fill=xfr_fill)
                        b = cv.create_oval(x+2, y-6, x+12, y+6, outline=text_color, width=2, fill=xfr_fill)
                    else:
                        a = cv.create_oval(x-6, y-12, x+6, y-2, outline=text_color, width=2, fill=xfr_fill)
                        b = cv.create_oval(x-6, y+2, x+6, y+12, outline=text_color, width=2, fill=xfr_fill)
                    tag_and_meta(a, 'Transformer', f'xfmr:{k}', name)
                    tag_and_meta(b, 'Transformer', f'xfmr:{k}', name)
                    hit = cv.create_rectangle(x-14, y-14, x+14, y+14, outline='', fill='')
                    tag_and_meta(hit, 'Transformer', f'xfmr:{k}', name)
                else:
                    closed = d.get('closed', True)
                    if horiz:
                        if closed:
                            sitem = cv.create_line(x-10, y, x+10, y, fill=switch_fill, width=3)
                        else:
                            cv.create_line(x-10, y, x-2, y, fill=switch_fill, width=3)
                            cv.create_line(x+2, y,  x+10, y, fill=switch_fill, width=3)
                            sitem = cv.create_line(x-2, y-6, x+2, y+6, fill=switch_fill, width=2)
                    else:
                        if closed:
                            sitem = cv.create_line(x, y-10, x, y+10, fill=switch_fill, width=3)
                        else:
                            cv.create_line(x, y-10, x, y-2, fill=switch_fill, width=3)
                            cv.create_line(x, y+2,  x, y+10, fill=switch_fill, width=3)
                            sitem = cv.create_line(x-6, y-2, x+6, y+2, fill=switch_fill, width=2)
                    tag_and_meta(sitem, 'Switch', f'switch:{k}', name)
                    hit = cv.create_rectangle(x-14, y-14, x+14, y+14, outline='', fill='')
                    tag_and_meta(hit, 'Switch', f'switch:{k}', name)

        # Draw buses
        node_fill = bus_fill
        node_stroke = text_color
        for n, (x, y) in coords_px.items():
            r = BUS_R
            item = cv.create_oval(x - r, y - r, x + r, y + r, fill=node_fill, outline=node_stroke, width=1.5)
            tag_and_meta(item, 'Bus', f'bus:{n}', n)
            # Invisible hit box
            hit = cv.create_rectangle(x-12, y-12, x+12, y+12, outline='', fill='')
            tag_and_meta(hit, 'Bus', f'bus:{n}', n)

        # Draw bus-attached devices (loads/shunts) and sources
        for b, cnt in bus_loads.items():
            if b not in coords_px:
                continue
            x, y = coords_px[b]
            for i in range(cnt):
                off = 12 + i * 10; s = 6
                item = cv.create_rectangle(x + off - s, y - s, x + off + s, y + s, outline=text_color, fill=load_fill, width=1)
                tag_and_meta(item, 'Load', f'load:{b}:{i+1}', f'Load {i+1} @ {b}')
        for b, cnt in bus_shunts.items():
            if b not in coords_px:
                continue
            x, y = coords_px[b]
            for i in range(cnt):
                base = 12 + i * 12
                top = y - base - 5; bot = y - base + 5
                cxm = x - 5; cxp = x + 5
                l1 = cv.create_line(cxm, top, cxm, bot, fill=shunt_fill, width=3)
                l2 = cv.create_line(cxp, top, cxp, bot, fill=shunt_fill, width=3)
                l3 = cv.create_line(cxm, bot, cxp, bot, fill=edge_color, width=1)
                for item in (l1, l2, l3):
                    tag_and_meta(item, 'Shunt', f'shunt:{b}:{i+1}', f'Shunt {i+1} @ {b}')
        for b in bus_sources:
            if b not in coords_px:
                continue
            x, y = coords_px[b]
            r = SRC_R
            circ = cv.create_oval(x - r, y - r, x + r, y + r, outline=text_color, fill=source_fill, width=2)
            bolt = cv.create_polygon(x, y-5, x+3, y, x+0, y+1, x+3, y+6, x-3, y+1, x+0, y, fill=text_color, outline='')
            tag_and_meta(circ, 'VoltageSource', f'source:{b}', f'Source @ {b}')
            tag_and_meta(bolt, 'VoltageSource', f'source:{b}', f'Source @ {b}')

        # Centering and scrollregion
        try:
            bbox = cv.bbox('all')
            if bbox:
                x1, y1, x2, y2 = bbox
                cv.configure(scrollregion=(x1 - 20, y1 - 20, x2 + 20, y2 + 20))
        except Exception:
            pass

        # Bind clicks once
        try:
            cv.tag_bind('bus', '<Button-1>', self._on_map_item_click)
            cv.tag_bind('edge', '<Button-1>', self._on_map_item_click)
            cv.tag_bind('dev', '<Button-1>', self._on_map_item_click)
        except Exception:
            pass

    def _on_map_item_click(self, event):
        try:
            cv = self.island_map
            if cv is None:
                return
            items = cv.find_withtag('current')
            if not items:
                return
            it = items[0]
            tags = cv.gettags(it)
            info = None
            for t in tags:
                if t in getattr(self, '_map_meta', {}):
                    info = self._map_meta.get(t)
                    break
            if info is None:
                # Edge/device tags also carry bus:/edge: prefixes
                for t in tags:
                    if t.startswith('bus:') or t.startswith('edge:'):
                        info = {"id": t.split(':',1)[1], "type": 'item'}
                        break
            # Clear previous tooltip
            try:
                if getattr(self, '_tooltip_item', None):
                    cv.delete(self._tooltip_item)
            except Exception:
                pass
            if info is None:
                return
            x = cv.canvasx(event.x); y = cv.canvasy(event.y)
            text = f"{info.get('type','item')}: {info.get('id','')}"
            self._tooltip_item = cv.create_text(x + 12, y - 12, text=text, anchor='nw', fill=self.COL.get('TEXT', '#334155'), font=(self.UI_FONT, max(10, self.UI_SIZE)))
        except Exception:
            pass

    # Unified OBJ interactions (orthogonal renderer)
    def _show_callout(self, x: float, y: float, text: str) -> None:
        try:
            cv = self.island_map
            if cv is None:
                return
            pad = 4
            # remove old
            if hasattr(self, '_callout'):
                box = self._callout.get('box'); txt = self._callout.get('txt')
                try:
                    if box: cv.delete(box)
                    if txt: cv.delete(txt)
                except Exception:
                    pass
            x2 = min(x + 180, max(8, (cv.winfo_width() or 0) - 8))
            y2 = max(y - 8, 8)
            box = cv.create_rectangle(x2, y2-16, x2+160, y2+6, fill='#fffffe', outline='#888', width=1)
            txt = cv.create_text(x2+pad, y2-5, text=text, anchor='w', fill=self.COL.get('TEXT', '#222'), font=(self.UI_FONT, max(9, self.UI_SIZE-1)))
            self._callout = {'box': box, 'txt': txt}
        except Exception:
            pass

    def _hide_callout(self) -> None:
        try:
            cv = self.island_map
            if cv is None:
                return
            if hasattr(self, '_callout'):
                box = self._callout.get('box'); txt = self._callout.get('txt')
                try:
                    if box: cv.delete(box)
                    if txt: cv.delete(txt)
                except Exception:
                    pass
                delattr(self, '_callout')
        except Exception:
            pass

    def _highlight_group(self, item_id: int) -> None:
        try:
            cv = self.island_map
            if cv is None:
                return
            text_color = self.COL.get('TEXT', '#222')
            select_outline = self.COL.get('SELECT', '#6a0dad')
            # clear prior
            if hasattr(self, '_sel'):
                for i in getattr(self, '_sel'):
                    try:
                        cv.itemconfig(i, width=1, outline=text_color)
                    except Exception:
                        pass
            tags = cv.gettags(item_id)
            id_tag = next((t for t in tags if t.startswith('id:')), None)
            if not id_tag:
                return
            group = cv.find_withtag(id_tag)
            for i in group:
                try:
                    cv.itemconfig(i, width=3, outline=select_outline)
                except Exception:
                    pass
            self._sel = group
        except Exception:
            pass

    def _on_obj_click(self, ev):
        try:
            cv = self.island_map
            if cv is None:
                return
            items = cv.find_withtag('current')
            if not items:
                return
            it = items[0]
            md = getattr(self, 'meta', {}).get(it)
            if not md:
                # try to resolve via same id tag
                tags = cv.gettags(it)
                id_tag = next((t for t in tags if t.startswith('id:')), None)
                if id_tag:
                    grp = cv.find_withtag(id_tag)
                    md = next((getattr(self, 'meta', {}).get(g) for g in grp if getattr(self, 'meta', {}).get(g)), None)
            if not md:
                return
            self._highlight_group(it)
            self._show_callout(ev.x, ev.y, f"{md.get('name', md.get('id',''))} ({md.get('type','')})")
        except Exception:
            pass

    def _on_obj_hover_enter(self, ev):
        try:
            cv = self.island_map
            if cv is None:
                return
            items = cv.find_withtag('current')
            if items:
                select_outline = self.COL.get('SELECT', '#6a0dad')
                try:
                    cv.itemconfig(items[0], width=3, outline=select_outline)
                except Exception:
                    pass
        except Exception:
            pass

    def _on_obj_hover_leave(self, ev):
        try:
            cv = self.island_map
            if cv is None:
                return
            items = cv.find_withtag('current')
            if items:
                text_color = self.COL.get('TEXT', '#222')
                try:
                    cv.itemconfig(items[0], width=1, outline=text_color)
                except Exception:
                    pass
            self._hide_callout()
        except Exception:
            pass

# ---- Entrypoint --------------------------------------------------------------
def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
