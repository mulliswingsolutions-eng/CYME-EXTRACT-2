# Modules/Pins.py
from __future__ import annotations
from pathlib import Path
import xml.etree.ElementTree as ET
from typing import Dict, List, Set, Tuple

from Modules.General import safe_name
from Modules.IslandFilter import should_comment_bus, should_comment_branch

PHASES = ("A", "B", "C")
SUFFIX = {"A": "_a", "B": "_b", "C": "_c"}


# ---------- Parsing helpers (all derived from the file) ----------

def _bus_phases(root: ET.Element) -> Dict[str, Set[str]]:
    """Collect phases present at each bus by scanning all Sections (sanitized)."""
    bus_ph: Dict[str, Set[str]] = {}
    for sec in root.findall(".//Section"):
        ph = (sec.findtext("Phase") or "ABC").upper()
        if not any(p in PHASES for p in ph):
            ph = "ABC"
        for tag in ("FromNodeID", "ToNodeID"):
            bus_raw = (sec.findtext(tag) or "").strip()
            if not bus_raw:
                continue
            bus = safe_name(bus_raw)
            bus_ph.setdefault(bus, set()).update([p for p in PHASES if p in ph])
    return bus_ph


def _loads_by_bus(root: ET.Element) -> Dict[str, int]:
    """Return bus -> number_of_phases_with_values (for PQ pins), using sanitized bus names."""
    out: Dict[str, int] = {}
    for sec in root.findall(".//Section"):
        spot = sec.find(".//Devices/SpotLoad")
        if spot is None:
            continue
        bus = safe_name((sec.findtext("FromNodeID") or "").strip())
        if not bus:
            continue
        phases = set()
        for val in spot.findall(".//CustomerLoadValue"):
            p = (val.findtext("Phase") or "").strip().upper()
            if p in PHASES:
                phases.add(p)
        out[bus] = max(out.get(bus, 0), len(phases) or 1)
    return out


def _shunt_buses(root: ET.Element) -> Set[str]:
    """Set of buses that host shunt capacitors (sanitized)."""
    out: Set[str] = set()
    for sec in root.findall(".//Section"):
        if sec.find(".//Devices/ShuntCapacitor") is not None:
            bus = safe_name((sec.findtext("FromNodeID") or "").strip())
            if bus:
                out.add(bus)
    return out


def _xfmr_secondary_buses(root: ET.Element) -> Set[str]:
    """
    Identify transformer sections and mark the *secondary* bus (sanitized).
    We infer the secondary as the node opposite NormalFeedingNodeID
    (falling back to ToNodeID).
    """
    secondaries: Set[str] = set()
    for sec in root.findall(".//Section"):
        xf = sec.find(".//Devices/Transformer")
        if xf is None:
            continue
        from_bus = safe_name((sec.findtext("FromNodeID") or "").strip())
        to_bus   = safe_name((sec.findtext("ToNodeID") or "").strip())
        normal   = safe_name((xf.findtext("NormalFeedingNodeID") or "").strip())
        if normal and normal == from_bus and to_bus:
            second = to_bus
        elif normal and normal == to_bus and from_bus:
            second = from_bus
        else:
            second = to_bus or from_bus
        if second:
            secondaries.add(second)
    return secondaries


def _line_sections(root: ET.Element) -> List[Tuple[str, str, int]]:
    """
    Return list of (from_bus, to_bus, nphases) for line-like devices (sanitized).
    Includes overhead and underground variants.
    """
    out: List[Tuple[str, str, int]] = []
    for sec in root.findall(".//Section"):
        if not (sec.find(".//Devices/OverheadLineUnbalanced") is not None or
                sec.find(".//Devices/OverheadByPhase") is not None or
                sec.find(".//Devices/OverheadLine") is not None or
                sec.find(".//Devices/Underground") is not None or
                sec.find(".//Devices/UndergroundCable") is not None):
            continue
        fb = safe_name((sec.findtext("FromNodeID") or "").strip())
        tb = safe_name((sec.findtext("ToNodeID") or "").strip())
        ph = (sec.findtext("Phase") or "ABC").upper()
        nph = sum(1 for p in PHASES if p in ph) or 3
        if fb and tb:
            out.append((fb, tb, nph))
    return out


def _transformer_pairs(root: ET.Element) -> List[Tuple[str, str, int]]:
    """Return list of (from_bus, to_bus, nphases) for transformers (for tap pins), sanitized."""
    out: List[Tuple[str, str, int]] = []
    for sec in root.findall(".//Section"):
        xf = sec.find(".//Devices/Transformer")
        if xf is None:
            continue
        fb = safe_name((sec.findtext("FromNodeID") or "").strip())
        tb = safe_name((sec.findtext("ToNodeID") or "").strip())
        ph = (sec.findtext("Phase") or "ABC").upper()
        nph = sum(1 for p in PHASES if p in ph) or 3
        if fb and tb:
            out.append((fb, tb, nph))
    return out


# ---------- Selection logic (no hardcoding) ----------

def _voltage_bus_set(root: ET.Element) -> Set[str]:
    """
    Build the set of buses for V_abs/V_ang pins (sanitized):
      - buses with SpotLoad(s)
      - buses with ShuntCapacitor(s)
      - transformer secondary buses
      - plus one-hop neighbors of any of the above (via a line)
    """
    loads = set(_loads_by_bus(root).keys())
    shunts = _shunt_buses(root)
    xsec = _xfmr_secondary_buses(root)

    base = loads | shunts | xsec

    # expand by one hop along line sections
    lines = _line_sections(root)
    neighbors: Dict[str, Set[str]] = {}
    for a, b, _ in lines:
        neighbors.setdefault(a, set()).add(b)
        neighbors.setdefault(b, set()).add(a)

    expanded = set(base)
    for b in list(base):
        expanded.update(neighbors.get(b, set()))

    return expanded


# ---------- Sheet writer ----------

def write_pins_sheet(xw, input_path: Path) -> None:
    """
    Create the 'Pins' sheet with rows:
      outgoing  V_abs   <bus_phase>/Vmag ...
      outgoing  V_ang   <bus_phase>/Vang ...
      outgoing  I_abs   LN_from_to/ImagFrom{1..n} ...
      outgoing  I_ang   LN_from_to/IangFrom{1..n} ...
      outgoing  PQ_ld   LD_<bus>/P# LD_<bus>/Q# ...
      incoming  Trans_tap TR1_from_to/tap_# ...
      incoming  PQ_ld   LD_<bus>/P# LD_<bus>/Q# ...

    Protections:
      - All bus/ID strings are sanitized via safe_name (so '-' -> '__', etc.).
      - **Island policy** drives inclusion:
          * If an active island is chosen â†’ include only that island.
          * If none chosen â†’ include only islands with a voltage source.
    """
    root = ET.fromstring(input_path.read_text(encoding="utf-8", errors="ignore"))

    # Discovery (all sanitized by helpers above)
    bus_ph = _bus_phases(root)                  # bus -> phases present
    v_buses = _voltage_bus_set(root)            # selected buses for voltage pins
    lines = _line_sections(root)                # all line sections (for currents)
    load_phase_counts = _loads_by_bus(root)     # bus -> number of load phases
    xf_pairs = _transformer_pairs(root)         # (from,to,nph)

    # ---------- Filter by island policy ----------
    # Buses
    v_buses = {b for b in v_buses if not should_comment_bus(b)}
    bus_ph = {b: phs for b, phs in bus_ph.items() if not should_comment_bus(b)}
    load_phase_counts = {b: n for b, n in load_phase_counts.items() if not should_comment_bus(b)}

    # Branch-like objects (lines / transformer pairs)
    lines = [(fb, tb, nph) for (fb, tb, nph) in lines if not should_comment_branch(fb, tb)]
    xf_pairs = [(fb, tb, nph) for (fb, tb, nph) in xf_pairs if not should_comment_branch(fb, tb)]

    # Writer setup
    wb = xw.book
    ws = wb.add_worksheet("Pins")
    xw.sheets["Pins"] = ws

    hdr = wb.add_format({"bold": True})
    ws.set_column(0, 0, 10)   # direction
    ws.set_column(1, 1, 12)   # group
    ws.set_column(2, 200, 24)

    r = 0
    def row(items):  # simple row write
        nonlocal r
        ws.write_row(r, 0, items)
        r += 1

    # ---- outgoing: V_abs ----
    vabs = ["//outgoing", "V_abs"]
    for b in sorted(v_buses):
        for p in sorted(bus_ph.get(b, set(PHASES)), key=lambda x: "ABC".index(x)):
            vabs.append(f"{b}{SUFFIX[p]}/Vmag")
    row(vabs)

    # ---- outgoing: V_ang ----
    vang = ["//outgoing", "V_ang"]
    for b in sorted(v_buses):
        for p in sorted(bus_ph.get(b, set(PHASES)), key=lambda x: "ABC".index(x)):
            vang.append(f"{b}{SUFFIX[p]}/Vang")
    row(vang)

    # ---- outgoing: I_abs (From side currents) ----
    iabs = ["//outgoing", "I_abs"]
    for fb, tb, nph in sorted(lines):
        tag = f"LN_{fb}_{tb}"
        for idx in range(1, nph + 1):
            iabs.append(f"{tag}/ImagFrom{idx}")
    row(iabs)

    # ---- outgoing: I_ang (From side angles) ----
    iang = ["//outgoing", "I_ang"]
    for fb, tb, nph in sorted(lines):
        tag = f"LN_{fb}_{tb}"
        for idx in range(1, nph + 1):
            iang.append(f"{tag}/IangFrom{idx}")
    row(iang)

    # ---- outgoing: PQ_ld (loads at the selected buses) ----
    pq_out = ["//outgoing", "PQ_ld"]
    for b in sorted(v_buses):
        n = load_phase_counts.get(b, 0)
        if n <= 0:
            continue
        base = f"LD_{b}"
        for idx in range(1, n + 1):
            pq_out.append(f"{base}/P{idx}")
            pq_out.append(f"{base}/Q{idx}")
    row(pq_out)

    # ---- incoming: Trans_tap (all transformers with active endpoints) ----
    taps = ["//incoming", "Trans_tap"]
    for fb, tb, nph in sorted(xf_pairs):
        base = f"TR1_{fb}_{tb}"
        for idx in range(1, nph + 1):
            taps.append(f"{base}/tap_{idx}")
    row(taps)

    # ---- incoming: PQ_ld (match the same loads we exposed) ----
    pq_in = ["//incoming", "PQ_ld"]
    for b in sorted(v_buses):
        n = load_phase_counts.get(b, 0)
        if n <= 0:
            continue
        base = f"LD_{b}"
        for idx in range(1, n + 1):
            pq_in.append(f"{base}/P{idx}")
            pq_in.append(f"{base}/Q{idx}")
    row(pq_in)
