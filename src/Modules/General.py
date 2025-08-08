# Modules/General.py
from __future__ import annotations
from pathlib import Path
from typing import List, Tuple, Optional
import re
import xml.etree.ElementTree as ET

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
    Return ONLY Frequency and BaseMVA from the single <GlobalParameters> block.
    If the block or values are missing, return None for that value (no defaults here).
    """
    file_path = Path(file_path)
    text = file_path.read_text(encoding="utf-8", errors="ignore")

    # Extract *only* the <GlobalParameters>...</GlobalParameters> chunk
    m = GP_BLOCK_RE.search(text)
    if not m:
        return [("Frequency (Hz)", None), ("Power Base (MVA)", None)]

    gp_chunk = "<GlobalParameters>" + m.group(1) + "</GlobalParameters>"

    # Parse just that chunk
    gp = ET.fromstring(gp_chunk)  # gp.tag == 'GlobalParameters'
    freq = _to_float(gp.findtext("Frequency"))
    base_mva = _to_float(gp.findtext("BaseMVA"))

    return [("Frequency (Hz)", freq), ("Power Base (MVA)", base_mva)]
