# Modules/Switch.py
from __future__ import annotations
from pathlib import Path
import xml.etree.ElementTree as ET
from typing import List, Tuple

PHASES = ("A", "B", "C")
SUFFIX = {"A": "_a", "B": "_b", "C": "_c"}
DEVICE_TAGS = ("Switch", "Sectionalizer", "Breaker", "Fuse")


def _phase_tokens(s: str | None) -> List[str]:
    """Extract A/B/C that appear in s (case-insensitive)."""
    if not s:
        return []
    u = s.upper()
    return [p for p in PHASES if p in u]


def _device_id(dev: ET.Element, from_bus: str, to_bus: str) -> str:
    """Prefer DeviceNumber, then DeviceID, else synthesize."""
    did = (dev.findtext("DeviceNumber") or "").strip()
    if not did:
        did = (dev.findtext("DeviceID") or "").strip()
    if not did:
        did = f"SW_{from_bus}_{to_bus}"
    return did


def _rows_from_file(txt_path: Path) -> List[Tuple[str, str, str, int]]:
    """
    Returns rows of (From Bus, To Bus, ID, Status).
    Status: 1 if closed on that phase, else 0.

    Devices included: Switch, Sectionalizer, Breaker, Fuse.
    """
    root = ET.fromstring(txt_path.read_text(encoding="utf-8", errors="ignore"))
    rows: List[Tuple[str, str, str, int]] = []

    for sec in root.findall(".//Section"):
        from_bus = (sec.findtext("FromNodeID") or "").strip()
        to_bus   = (sec.findtext("ToNodeID") or "").strip()
        if not from_bus or not to_bus:
            continue

        # Section's declared phases (can be empty/None/invalid).
        sec_phase_text = sec.findtext("Phase")
        sec_phases = _phase_tokens(sec_phase_text)

        devs: List[ET.Element] = []
        for tag in DEVICE_TAGS:
            devs.extend(sec.findall(f".//Devices/{tag}"))

        for dev in devs:
            base_id = _device_id(dev, from_bus, to_bus)

            closed_phase_text = (dev.findtext("ClosedPhase") or "").strip()
            closed_set = set(_phase_tokens(closed_phase_text))

            # If section phases are missing/invalid, fall back to ClosedPhase; else ABC.
            phases = sec_phases if sec_phases else (list(closed_set) if closed_set else list(PHASES))

            normal_status = (dev.findtext("NormalStatus") or "").strip().lower()  # "closed"/"open"

            for p in phases:
                # ClosedPhase (if present) is authoritative per phase; otherwise use NormalStatus.
                if closed_set:
                    is_closed = (p in closed_set)
                else:
                    is_closed = (normal_status == "closed")

                fb = f"{from_bus}{SUFFIX[p]}"
                tb = f"{to_bus}{SUFFIX[p]}"
                rid = f"{base_id}{SUFFIX[p]}"
                rows.append((fb, tb, rid, 1 if is_closed else 0))

    rows.sort(key=lambda r: (r[0], r[1], r[2]))
    return rows


def write_switch_sheet(xw, input_path: Path) -> None:
    """
    Create the 'Switch' sheet with columns:
    From Bus | To Bus | ID | Status
    (Now includes Switches, Sectionalizers, Breakers, and Fuses.)
    """
    wb = xw.book
    ws = wb.add_worksheet("Switch")
    xw.sheets["Switch"] = ws

    # Formats
    header = wb.add_format({"bold": True})
    int0 = wb.add_format({"num_format": "0"})

    # Column widths
    ws.set_column(0, 0, 18)  # From Bus
    ws.set_column(1, 1, 18)  # To Bus
    ws.set_column(2, 2, 22)  # ID (DeviceNumber/DeviceID + _phase)
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
