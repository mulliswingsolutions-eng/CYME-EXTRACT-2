# Modules/Switch.py
from __future__ import annotations
from pathlib import Path
import re
import xml.etree.ElementTree as ET
from typing import List, Tuple
from Modules.General import safe_name

PHASES = ("A", "B", "C")
SUFFIX = {"A": "_a", "B": "_b", "C": "_c"}

# Native switch-like device tags we export as switches
DEVICE_TAGS = ("Switch", "Sectionalizer", "Breaker", "Fuse")

# Miscellaneous devices to TREAT AS series switches
# Includes RB (meter) and LA (lightning arrester).
MISC_AS_SWITCH_IDS = {"RB", "LA"}  # case-insensitive match on <DeviceID>

# For stripping phase suffixes when needed
_PHASE_SUFFIX_RE = re.compile(r"_(a|b|c)$")


def _phase_tokens(s: str | None) -> List[str]:
    if not s:
        return []
    u = s.upper()
    return [p for p in PHASES if p in u]


def _device_id(dev: ET.Element, from_bus_san: str, to_bus_san: str) -> str:
    did = (dev.findtext("DeviceNumber") or "").strip()
    if not did:
        did = (dev.findtext("DeviceID") or "").strip()
    did = safe_name(did)
    if not did:
        did = safe_name(f"SW_{from_bus_san}_{to_bus_san}")
    return did


def _bool_from_text(s: str | None, default: bool | None = None) -> bool | None:
    if s is None:
        return default
    t = s.strip().lower()
    if t in {"1", "true", "yes", "y", "connected", "closed"}:
        return True
    if t in {"0", "false", "no", "n", "disconnected", "open"}:
        return False
    return default


def _keep_device(dev: ET.Element, dev_type: str) -> bool:
    """
    Location-based filter for native switch-like devices:
      - Keep Location="Middle" or missing.
      - For Location in {"From","To"}:
          * Breaker -> keep only if Restriction == 0
          * Others  -> keep regardless of Restriction
    """
    loc = (dev.findtext("Location") or "").strip().lower()
    if not loc or loc == "middle":
        return True

    if loc in {"from", "to"}:
        if dev_type == "Breaker":
            restr_txt = (dev.findtext("Restriction") or "").strip()
            is_restricted = (restr_txt not in ("", "0", "false", "False"))
            return not is_restricted
        return True

    if dev_type == "Breaker":
        restr_txt = (dev.findtext("Restriction") or "").strip()
        is_restricted = (restr_txt not in ("", "0", "false", "False"))
        return not is_restricted
    return True


def _active_bus_bases_from_bus_sheet(input_path: Path) -> set[str]:
    """
    Consider a bus active if it appears on the Bus sheet and does NOT start with '//'.
    Return base names (without trailing _a/_b/_c).
    """
    # Lazy import to avoid circular imports
    from Modules.Bus import extract_bus_data

    bases: set[str] = set()
    for row in extract_bus_data(input_path):
        bus = str(row.get("Bus", "")).strip()
        if not bus or bus.startswith("//"):
            continue
        base = _PHASE_SUFFIX_RE.sub("", bus)
        bases.add(base)
    return bases


def _rows_from_file(txt_path: Path) -> List[Tuple[str, str, str, int]]:
    """
    Returns rows of (From Bus, To Bus, ID, Status).
    Status: 1 if closed on that phase, else 0.

    Includes:
      - Switch/Sectionalizer/Breaker/Fuse (with location filtering)
      - Miscellaneous with DeviceID in MISC_AS_SWITCH_IDS
        (treated as series switches placed between FromNodeID and ToNodeID)

    NOTE: If either endpoint bus base is NOT active on the Bus sheet,
          the row is commented by prefixing '//' to the **From Bus** cell
          (leftmost column). The ID remains unprefixed.
    """
    root = ET.fromstring(txt_path.read_text(encoding="utf-8", errors="ignore"))
    rows: List[Tuple[str, str, str, int]] = []

    # Build the set of active bus bases from the Bus sheet
    active_bases = _active_bus_bases_from_bus_sheet(txt_path)

    for sec in root.findall(".//Section"):
        from_bus_raw = (sec.findtext("FromNodeID") or "").strip()
        to_bus_raw   = (sec.findtext("ToNodeID") or "").strip()
        if not from_bus_raw or not to_bus_raw:
            continue

        from_bus = safe_name(from_bus_raw)  # base (no phase suffix)
        to_bus   = safe_name(to_bus_raw)    # base (no phase suffix)

        sec_phases = _phase_tokens(sec.findtext("Phase"))

        # Should this row be commented?
        def _comment_row() -> bool:
            return not ((from_bus in active_bases) and (to_bus in active_bases))

        # --- native switch-like devices ---
        for tag in DEVICE_TAGS:
            for dev in sec.findall(f".//Devices/{tag}"):
                if not _keep_device(dev, tag):
                    continue

                base_id = _device_id(dev, from_bus, to_bus)

                closed_phase_text = (dev.findtext("ClosedPhase") or "").strip()
                closed_set = set(_phase_tokens(closed_phase_text))  # "None" -> empty set

                phases = sec_phases if sec_phases else (list(closed_set) if closed_set else list(PHASES))
                normal_status = _bool_from_text(dev.findtext("NormalStatus"), default=None)

                for p in phases:
                    is_closed = (p in closed_set) if closed_set else (bool(normal_status) if normal_status is not None else False)

                    fb = f"{from_bus}{SUFFIX[p]}"
                    tb = f"{to_bus}{SUFFIX[p]}"
                    rid = f"{base_id}{SUFFIX[p]}"

                    # Put '//' in the **From Bus** cell when commenting
                    fb_out = f"//{fb}" if _comment_row() else fb
                    rows.append((fb_out, tb, rid, 1 if is_closed else 0))

        # --- Miscellaneous → treat selected DeviceID codes as series switches ---
        for dev in sec.findall(".//Devices/Miscellaneous"):
            dev_code = ((dev.findtext("DeviceID") or "").strip().upper())
            if dev_code not in MISC_AS_SWITCH_IDS:
                continue

            # Consider connected==closed; default to closed for these inline devices
            conn = _bool_from_text(dev.findtext("ConnectionStatus"), default=True)
            is_closed = True if conn is None else bool(conn)

            base_id = _device_id(dev, from_bus, to_bus)
            phases = sec_phases if sec_phases else list(PHASES)

            for p in phases:
                fb = f"{from_bus}{SUFFIX[p]}"
                tb = f"{to_bus}{SUFFIX[p]}"
                rid = f"{base_id}{SUFFIX[p]}"
                fb_out = f"//{fb}" if _comment_row() else fb
                rows.append((fb_out, tb, rid, 1 if is_closed else 0))

    # De-dup identical rows
    rows = sorted(set(rows), key=lambda r: (r[0], r[1], r[2], r[3]))
    return rows


def write_switch_sheet(xw, input_path: Path) -> None:
    """
    Create the 'Switch' sheet with columns:
    From Bus | To Bus | ID | Status

    Includes Switch/Sectionalizer/Breaker/Fuse (filtered) and
    Miscellaneous with DeviceID in MISC_AS_SWITCH_IDS as closed series switches.

    Rows that should be disabled are commented by prefixing '//' on the **From Bus** cell.
    """
    wb = xw.book
    ws = wb.add_worksheet("Switch")
    xw.sheets["Switch"] = ws

    header = wb.add_format({"bold": True})
    int0 = wb.add_format({"num_format": "0"})

    ws.set_column(0, 0, 18)  # From Bus
    ws.set_column(1, 1, 18)  # To Bus
    ws.set_column(2, 2, 28)  # ID
    ws.set_column(3, 3, 8)   # Status

    ws.write_row(0, 0, ["From Bus", "To Bus", "ID", "Status"], header)

    rows = _rows_from_file(input_path)
    for r, (fb, tb, rid, status) in enumerate(rows, start=1):
        ws.write(r, 0, fb)
        ws.write(r, 1, tb)
        ws.write(r, 2, rid)
        ws.write_number(r, 3, status, int0)
