# Modules/General.py
from __future__ import annotations
from pathlib import Path
from typing import List, Tuple, Optional
import re
import xml.etree.ElementTree as ET
import re
from typing import Optional, Any, Dict

# ================================
# Editable constants
EXCEL_FILE_VERSION = "v2.0"
SYSTEM_NAME = "Distribution"
# ================================

_GP_BLOCK_RE = re.compile(
    r"<GlobalParameters\b[^>]*>(.*?)</GlobalParameters>",
    re.DOTALL | re.IGNORECASE,
)

# --------------- Global Safe Name ---------------
_INVALID_RE = re.compile(r"[^A-Za-z0-9_]+")
def safe_name(s: Optional[str]) -> str:
    """
    Sanitize identifiers:
      - '-'  -> '__'   (so '6124-19' becomes '6124__19')
      - all other non [A-Za-z0-9_] -> '_'
      - collapse runs of 3+ underscores to a single '_' (preserves '__')
      - trim leading/trailing underscores
    """
    s = (s or "").strip()
    if not s:
        return ""
    s = s.replace("-", "__")               # special rule
    s = _INVALID_RE.sub("_", s)            # everything else -> '_'
    s = re.sub(r"_{3,}", "_", s)           # collapse 3+ underscores, keep '__'
    return s.strip("_")

# --------------- Global Island Context ---------------
# Writers will read this to decide commenting and SLACK assignment.
_ISLAND_CTX: Dict[str, Any] = {}

def set_island_context(ctx: Dict[str, Any]) -> None:
    global _ISLAND_CTX
    _ISLAND_CTX = ctx or {}

def get_island_context() -> Dict[str, Any]:
    return _ISLAND_CTX

# --------------- General Page ---------------
def _to_float(x: Optional[str]) -> Optional[float]:
    if x is None:
        return None
    try:
        return float(x.strip())
    except Exception:
        return None

def _parse_general(file_path: Path) -> list[tuple[str, Optional[float] | str]]:
    """
    Parse <GlobalParameters> from the CYME text and return the 2-column rows
    we want to write on the 'General' sheet.
    """
    text = Path(file_path).read_text(encoding="utf-8", errors="ignore")

    m = _GP_BLOCK_RE.search(text)
    if not m:
        freq = None
        base_mva = None
    else:
        gp_chunk = "<GlobalParameters>" + m.group(1) + "</GlobalParameters>"
        gp = ET.fromstring(gp_chunk)
        freq = _to_float(gp.findtext("Frequency"))
        base_mva = _to_float(gp.findtext("BaseMVA"))

    return [
        ("Excel file version", EXCEL_FILE_VERSION),
        ("Name", SYSTEM_NAME),
        ("Frequency (Hz)", freq),
        ("Power Base (MVA)", base_mva),
    ]

# --- Backward-compat helper (keeps your old call site working) ---
def get_general(file_path: str | Path) -> List[Tuple[str, Optional[float] | str]]:
    return _parse_general(Path(file_path))

# --- New unified API, same as the other modules ---
def write_general_sheet(xw, input_path: Path) -> None:
    """
    Create the 'General' sheet using the same xlsxwriter pattern as other pages.
    """
    rows = _parse_general(Path(input_path))

    wb = xw.book
    ws = wb.add_worksheet("General")
    # (Optional) store handle for consistency with your other modules
    try:
        xw.sheets["General"] = ws  # pandas.ExcelWriter keeps this dict
    except Exception:
        pass

    # Formats
    key_fmt = wb.add_format({"bold": True})
    num_fmt = wb.add_format({"num_format": "0.00"})

    # Column widths
    ws.set_column(0, 0, 24)  # labels
    ws.set_column(1, 1, 18)  # values

    # Write rows
    r = 0
    for label, value in rows:
        ws.write(r, 0, label, key_fmt)
        if isinstance(value, (int, float)) and value is not None:
            ws.write_number(r, 1, float(value), num_fmt)
        elif value is None:
            ws.write(r, 1, "")
        else:
            ws.write(r, 1, str(value))
        r += 1
