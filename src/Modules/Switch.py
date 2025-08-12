# Modules/Switch.py
from __future__ import annotations
from pathlib import Path
import xml.etree.ElementTree as ET
from typing import List, Tuple

PHASES = ("A", "B", "C")
SUFFIX = {"A": "_a", "B": "_b", "C": "_c"}


def _phase_list(section_phase: str | None) -> List[str]:
    s = (section_phase or "ABC").upper()
    return [p for p in PHASES if p in s]


def _rows_from_file(txt_path: Path) -> List[Tuple[str, str, str, int]]:
    """
    Returns rows of (From Bus, To Bus, ID, Status)
    Status: 1 if closed on that phase, else 0.
    """
    root = ET.fromstring(txt_path.read_text(encoding="utf-8", errors="ignore"))
    rows: List[Tuple[str, str, str, int]] = []

    for sec in root.findall(".//Section"):
        sw = sec.find(".//Devices/Switch")
        if sw is None:
            continue

        from_bus = (sec.findtext("FromNodeID") or "").strip()
        to_bus = (sec.findtext("ToNodeID") or "").strip()
        phase_tag = sec.findtext("Phase")

        normal_status = (sw.findtext("NormalStatus") or "").strip().lower()   # "closed" / "open"
        closed_phase = (sw.findtext("ClosedPhase") or "").strip().upper()     # e.g., "ABC", "AB", "A", ""

        phases_in_section = _phase_list(phase_tag)
        closed_set = set(p for p in PHASES if p in closed_phase)

        for p in phases_in_section:
            # Determine per-phase status:
            # If ClosedPhase is provided, trust it; otherwise fall back to NormalStatus.
            if closed_phase:
                is_closed = p in closed_set
            else:
                is_closed = (normal_status == "closed")

            fb = f"{from_bus}{SUFFIX[p]}"
            tb = f"{to_bus}{SUFFIX[p]}"
            rid = f"SW_{from_bus}_{to_bus}{SUFFIX[p]}"
            rows.append((fb, tb, rid, 1 if is_closed else 0))

    # Stable sort for nice ordering
    rows.sort(key=lambda r: (r[0], r[1], r[2]))
    return rows


def write_switch_sheet(xw, input_path: Path) -> None:
    """
    Create the 'Switch' sheet with columns:
    From Bus | To Bus | ID | Status
    """
    wb = xw.book
    ws = wb.add_worksheet("Switch")
    xw.sheets["Switch"] = ws

    # Formats
    header = wb.add_format({"bold": True})
    int0 = wb.add_format({"num_format": "0"})

    # Column widths
    ws.set_column(0, 0, 14)  # From Bus
    ws.set_column(1, 1, 14)  # To Bus
    ws.set_column(2, 2, 24)  # ID
    ws.set_column(3, 3, 8)   # Status

    # Header
    ws.write_row(0, 0, ["From Bus", "To Bus", "ID", "Status"], header)

    # Data
    rows = _rows_from_file(input_path)
    r = 1
    for fb, tb, rid, status in rows:
        ws.write(r, 0, fb)
        ws.write(r, 1, tb)
        ws.write(r, 2, rid)
        ws.write_number(r, 3, status, int0)
        r += 1
