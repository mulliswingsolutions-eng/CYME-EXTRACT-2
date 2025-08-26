# Modules/Switch.py
from __future__ import annotations
from pathlib import Path
import re
import xml.etree.ElementTree as ET
from typing import List, Tuple

PHASES = ("A", "B", "C")
SUFFIX = {"A": "_a", "B": "_b", "C": "_c"}
DEVICE_TAGS = ("Switch", "Sectionalizer", "Breaker", "Fuse")

# --- NEW: sanitize identifiers (allow only [A-Za-z0-9_]) ---
_SAFE_RE = re.compile(r"[^A-Za-z0-9_]+")
def _safe_name(s: str | None) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    s = _SAFE_RE.sub("_", s)       # replace disallowed chars with "_"
    s = re.sub(r"_+", "_", s)      # collapse runs of "_"
    return s.strip("_")            # trim leading/trailing "_"


def _phase_tokens(s: str | None) -> List[str]:
    """Extract A/B/C that appear in s (case-insensitive)."""
    if not s:
        return []
    u = s.upper()
    return [p for p in PHASES if p in u]


def _device_id(dev: ET.Element, from_bus_san: str, to_bus_san: str) -> str:
    """Prefer DeviceNumber, then DeviceID, else synthesize (all sanitized)."""
    did = (dev.findtext("DeviceNumber") or "").strip()
    if not did:
        did = (dev.findtext("DeviceID") or "").strip()
    did = _safe_name(did)
    if not did:
        did = _safe_name(f"SW_{from_bus_san}_{to_bus_san}")
    return did


def _keep_device(dev: ET.Element, dev_type: str) -> bool:
    """
    Location-based filter:
      - Always keep Location="Middle" or missing.
      - For edge devices (Location="From"/"To"):
          * Breaker  -> keep only if Restriction == 0
          * Others   -> keep regardless of Restriction
    """
    loc = (dev.findtext("Location") or "").strip().lower()
    if not loc or loc == "middle":
        return True

    if loc in ("from", "to"):
        if dev_type == "Breaker":
            restr_txt = (dev.findtext("Restriction") or "").strip()
            is_restricted = (restr_txt not in ("", "0", "false", "False"))
            return not is_restricted
        else:
            return True  # Switch / Sectionalizer / Fuse: include even if restricted

    # Unknown location values: treat like edges with the same logic
    if dev_type == "Breaker":
        restr_txt = (dev.findtext("Restriction") or "").strip()
        is_restricted = (restr_txt not in ("", "0", "false", "False"))
        return not is_restricted
    return True


def _rows_from_file(txt_path: Path) -> List[Tuple[str, str, str, int]]:
    """
    Returns rows of (From Bus, To Bus, ID, Status).
    Status: 1 if closed on that phase, else 0.

    Devices included: Switch, Sectionalizer, Breaker, Fuse.
    """
    root = ET.fromstring(txt_path.read_text(encoding="utf-8", errors="ignore"))
    rows: List[Tuple[str, str, str, int]] = []

    for sec in root.findall(".//Section"):
        from_bus_raw = (sec.findtext("FromNodeID") or "").strip()
        to_bus_raw   = (sec.findtext("ToNodeID") or "").strip()
        if not from_bus_raw or not to_bus_raw:
            continue

        # sanitize bus names
        from_bus = _safe_name(from_bus_raw)
        to_bus   = _safe_name(to_bus_raw)

        sec_phases = _phase_tokens(sec.findtext("Phase"))

        for tag in DEVICE_TAGS:
            for dev in sec.findall(f".//Devices/{tag}"):
                if not _keep_device(dev, tag):
                    continue

                base_id = _device_id(dev, from_bus, to_bus)

                closed_phase_text = (dev.findtext("ClosedPhase") or "").strip()
                closed_set = set(_phase_tokens(closed_phase_text))  # "None" -> empty set

                # Use section phases if available; else ClosedPhase if specified; else ABC
                phases = sec_phases if sec_phases else (list(closed_set) if closed_set else list(PHASES))

                normal_status = (dev.findtext("NormalStatus") or "").strip().lower()  # "closed"/"open"

                for p in phases:
                    # ClosedPhase (if present) wins; otherwise fall back to NormalStatus
                    is_closed = (p in closed_set) if closed_set else (normal_status == "closed")

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
    (Includes Switches, Sectionalizers, Breakers, and Fuses; device-aware location filtering.
     All bus/ID strings are sanitized to contain only letters, digits, and underscores.)
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
    ws.set_column(2, 2, 28)  # ID
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
