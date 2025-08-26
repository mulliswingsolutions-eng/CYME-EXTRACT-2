# Modules/Bus.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Set
import re
import xml.etree.ElementTree as ET

PHASES = ("A", "B", "C")


def _read_xml(path: Path) -> ET.Element:
    return ET.fromstring(path.read_text(encoding="utf-8", errors="ignore"))


# ---------- NEW: name sanitizer (removes '-' and any non [A-Za-z0-9_]) ----------
_SAFE_RE = re.compile(r"[^A-Za-z0-9_]+")

def _safe_name(s: str | None) -> str:
    """Return an identifier made of [A-Za-z0-9_] only (no hyphens)."""
    s = (s or "").strip()
    if not s:
        return ""
    s = _SAFE_RE.sub("_", s)      # replace disallowed chars (incl. '-') with _
    s = re.sub(r"_+", "_", s)     # collapse repeats
    return s.strip("_")           # trim leading/trailing _


def _phase_set(s: str | None) -> Set[str]:
    p = (s or "ABC").strip().upper()
    return {ch for ch in p if ch in PHASES} or set(PHASES)


def _has_any(sec: ET.Element, tags: List[str]) -> bool:
    devs = sec.find("./Devices")
    if devs is None:
        return False
    for t in tags:
        if devs.find(t) is not None:
            return True
    return False


def _local_pseudos(sec: ET.Element) -> Set[str]:
    """Identifiers local to this section only (used to detect local pseudo 'To' nodes)."""
    pseudos: Set[str] = set()
    sid = _safe_name(sec.findtext("./SectionID"))
    if sid:
        pseudos.add(sid)
    devs = sec.find("./Devices")
    if devs is not None:
        for dev in list(devs):
            dn = _safe_name(dev.findtext("DeviceNumber"))
            di = _safe_name(dev.findtext("DeviceID"))
            if dn:
                pseudos.add(dn)
            if di:
                pseudos.add(di)
    return pseudos


def _parse_bus_rows(input_path: Path) -> List[Dict]:
    """
    Build rows for the Bus sheet.

    Columns:
      Bus | Base Voltage (V) | Initial Vmag | Unit | Angle | Type
    """
    root = _read_xml(Path(input_path))

    # --- Slack/source info (sanitize node id) ---
    source_node = _safe_name(root.findtext(".//Sources/Source/SourceNodeID"))

    eq = root.find(
        ".//Sources/Source/EquivalentSourceModels/EquivalentSourceModel/EquivalentSource"
    )
    if eq is None:
        return []

    def _f(x: str | None, default: float = 0.0) -> float:
        try:
            return float(x) if x not in (None, "") else default
        except Exception:
            return default

    # Source LN kV and angles (definition-level values for Bus page)
    phase_v_kv = {
        "A": _f(eq.findtext("OperatingVoltage1")),
        "B": _f(eq.findtext("OperatingVoltage2")),
        "C": _f(eq.findtext("OperatingVoltage3")),
    }
    phase_ang = {
        "A": _f(eq.findtext("OperatingAngle1")),
        "B": _f(eq.findtext("OperatingAngle2")),
        "C": _f(eq.findtext("OperatingAngle3")),
    }

    # -------- PASS 1: scan all sections to determine real endpoints & pseudo candidates
    BRANCH_TAGS = [
        # Overhead / Underground lines (all flavors)
        "OverheadLine", "OverheadLineUnbalanced", "OverheadByPhase",
        "Underground", "UndergroundCable", "UndergroundCableUnbalanced", "UndergroundByPhase",
        # Transformers
        "Transformer",
        # Switching / protection devices (treat like branches)
        "Switch", "Fuse", "Recloser", "Breaker", "Sectionalizer", "Isolator",
    ]

    branch_endpoints: Set[str] = set()           # all From/To seen in branch sections (sanitized)
    all_from_nodes: Set[str] = set()             # all FromNodeID across sections (sanitized)
    device_terminal_candidates: Set[str] = set() # ToNodeID of SpotLoad/Shunt sections (sanitized)

    # For branch-local pseudo detection:
    from_count: Dict[str, int] = {}
    to_count_nonlocal: Dict[str, int] = {}       # times a node is To in a section where it's NOT local pseudo
    local_pseudo_to_candidates: Set[str] = set() # ToNodeIDs equal to local SectionID/DeviceNumber/DeviceID (sanitized)

    for sec in root.findall(".//Sections/Section"):
        f_raw = (sec.findtext("./FromNodeID") or "").strip()
        t_raw = (sec.findtext("./ToNodeID") or "").strip()
        f = _safe_name(f_raw)
        t = _safe_name(t_raw)

        if f:
            all_from_nodes.add(f)
            from_count[f] = from_count.get(f, 0) + 1

        ph = (sec.findtext("./Phase") or "ABC").strip().upper()

        has_branch = _has_any(sec, BRANCH_TAGS)
        has_spot = sec.find(".//Devices/SpotLoad") is not None
        has_shunt = (
            sec.find(".//Devices/ShuntCapacitor") is not None
            or sec.find(".//Devices/ShuntReactor") is not None
        )

        if has_branch:
            if f:
                branch_endpoints.add(f)
            if t:
                branch_endpoints.add(t)
            # local pseudo test for the To side of this branch
            if t:
                lp = _local_pseudos(sec)  # already sanitized inside
                if t in lp:
                    local_pseudo_to_candidates.add(t)
                else:
                    to_count_nonlocal[t] = to_count_nonlocal.get(t, 0) + 1

        if has_spot or has_shunt:
            # candidate true device terminals (we'll refine after the scan)
            if t:
                device_terminal_candidates.add(t)

    # Refined exclusions
    terminal_exclusions_1: Set[str] = {
        n for n in device_terminal_candidates if n not in branch_endpoints and n not in all_from_nodes
    }
    terminal_exclusions_2: Set[str] = {
        n for n in local_pseudo_to_candidates if from_count.get(n, 0) == 0 and to_count_nonlocal.get(n, 0) == 0
    }
    terminal_exclusions: Set[str] = terminal_exclusions_1 | terminal_exclusions_2

    # -------- PASS 2: build node -> phases map with refined exclusions
    def _is_real_bus(nid: str | None) -> bool:
        nid = _safe_name(nid)
        if not nid:
            return False
        if nid in terminal_exclusions:
            return False
        return True

    node_phases: Dict[str, Set[str]] = {}

    def _add(nid: str | None, pstr: str) -> None:
        nid = _safe_name(nid)
        if not _is_real_bus(nid):
            return
        s = node_phases.setdefault(nid, set())
        s |= _phase_set(pstr)

    for sec in root.findall(".//Sections/Section"):
        ph = (sec.findtext("./Phase") or "ABC").strip().upper()
        f = sec.findtext("./FromNodeID")
        t = sec.findtext("./ToNodeID")

        has_branch = _has_any(sec, BRANCH_TAGS)
        has_spot = sec.find(".//Devices/SpotLoad") is not None
        has_shunt = (
            sec.find(".//Devices/ShuntCapacitor") is not None
            or sec.find(".//Devices/ShuntReactor") is not None
        )

        if has_spot or has_shunt:
            # Device sections: include only the network side (FromNodeID)
            _add(f, ph)
            continue

        if has_branch:
            # Branch sections: include both ends (unless excluded by refined rules)
            _add(f, ph)
            _add(t, ph)
            continue

        # Fallback: conservative
        _add(f, ph)

    if source_node:
        node_phases.setdefault(source_node, set()).update(PHASES)

    # Emit rows
    def pkey(p: str) -> int:
        return PHASES.index(p)

    rows: List[Dict] = []
    for node in sorted(node_phases):
        for ph in sorted(node_phases[node], key=pkey):
            v_ln = phase_v_kv[ph] * 1000.0
            rows.append(
                {
                    "Bus": f"{node}_{ph.lower()}",
                    "Base Voltage (V)": v_ln,
                    "Initial Vmag": v_ln,
                    "Unit": "V",
                    "Angle": phase_ang[ph],
                    "Type": "SLACK" if node == source_node else "PQ",
                }
            )
    return rows


# Backward-compat API
def extract_bus_data(filepath: str | Path) -> List[Dict]:
    return _parse_bus_rows(Path(filepath))


# Unified writer API
def write_bus_sheet(xw, input_path: Path) -> None:
    rows = _parse_bus_rows(Path(input_path))

    wb = xw.book
    ws = wb.add_worksheet("Bus")
    xw.sheets["Bus"] = ws

    hdr = wb.add_format({"bold": True})
    num0 = wb.add_format({"num_format": "0"})
    num2 = wb.add_format({"num_format": "0.00"})

    ws.set_column(0, 0, 18)
    ws.set_column(1, 2, 18)
    ws.set_column(3, 3, 8)
    ws.set_column(4, 5, 10)

    ws.write_row(0, 0, ["Bus", "Base Voltage (V)", "Initial Vmag", "Unit", "Angle", "Type"], hdr)
    r = 1
    for row in rows:
        ws.write(r, 0, row["Bus"])
        ws.write_number(r, 1, float(row["Base Voltage (V)"]), num0)
        ws.write_number(r, 2, float(row["Initial Vmag"]), num0)
        ws.write(r, 3, row["Unit"])
        ws.write_number(r, 4, float(row["Angle"]), num2)
        ws.write(r, 5, row["Type"])
        r += 1
