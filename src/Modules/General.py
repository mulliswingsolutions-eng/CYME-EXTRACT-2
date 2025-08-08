# Modules/General.py
from __future__ import annotations
from pathlib import Path
from typing import List, Tuple, Optional
import re
import xml.etree.ElementTree as ET

# ================================
# Editable constants
EXCEL_FILE_VERSION = "v2.0"
SYSTEM_NAME = "Distribution"
# ================================

GP_BLOCK_RE = re.compile(
    r"<GlobalParameters\b[^>]*>(.*?)</GlobalParameters>",
    re.DOTALL | re.IGNORECASE,
)

def _to_float(x: Optional[str]) -> Optional[float]:
    if x is None:
        return None
    try:
        return float(x.strip())
    except Exception:
        return None

def get_general(file_path: str | Path) -> List[Tuple[str, Optional[float]]]:
    """
    Return:
    - Excel file version (hardcoded)
    - Name (hardcoded)
    - Frequency (from GlobalParameters)
    - Power Base (MVA) (from GlobalParameters)
    """
    file_path = Path(file_path)
    text = file_path.read_text(encoding="utf-8", errors="ignore")

    # Extract <GlobalParameters>...</GlobalParameters>
    m = GP_BLOCK_RE.search(text)
    if not m:
        freq = None
        base_mva = None
    else:
        gp_chunk = "<GlobalParameters>" + m.group(1) + "</GlobalParameters>"
        gp = ET.fromstring(gp_chunk)
        freq = _to_float(gp.findtext("Frequency"))
        base_mva = _to_float(gp.findtext("BaseMVA"))

    # Return data in desired order
    return [
        ("Excel file version", EXCEL_FILE_VERSION),
        ("Name", SYSTEM_NAME),
        ("Frequency (Hz)", freq),
        ("Power Base (MVA)", base_mva)
    ]
