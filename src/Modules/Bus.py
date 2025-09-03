# Modules/Bus.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Set, Tuple
import xml.etree.ElementTree as ET

from Modules.General import safe_name
from Modules.IslandFilter import should_comment_bus   # <-- use island policy here

PHASES = ("A", "B", "C")


def _read_xml(path: Path) -> ET.Element:
    return ET.fromstring(path.read_text(encoding="utf-8", errors="ignore"))


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
    sid = safe_name(sec.findtext("./SectionID"))
    if sid:
        pseudos.add(sid)
    devs = sec.find("./Devices")
    if devs is not None:
        for dev in list(devs):
            dn = safe_name(dev.findtext("DeviceNumber"))
            di = safe_name(dev.findtext("DeviceID"))
            if dn:
                pseudos.add(dn)
            if di:
                pseudos.add(di)
    return pseudos


def _gather_vs_page_sources_and_kvll(root: ET.Element) -> Tuple[Set[str], Dict[str, float]]:
    """
    Return (nodes, kvll_volts) for the sources that WILL SHOW on the Voltage Source page.
    Logic: Topo where NetworkType == 'Substation' and EquivalentMode != '1';
           source must have an EquivalentSource block.
    """
    nodes: Set[str] = set()
    kvll_map: Dict[str, float] = {}
    for topo in root.findall(".//Topo"):
        ntype = (topo.findtext("NetworkType") or "").strip().lower()
        eq_mode = (topo.findtext("EquivalentMode") or "").strip()
        if ntype != "substation" or eq_mode == "1":
            continue
        srcs = topo.find("./Sources")
        if srcs is None:
            continue
        for src in srcs.findall("./Source"):
            nid = safe_name(src.findtext("SourceNodeID"))
            eq = src.find("./EquivalentSourceModels/EquivalentSourceModel/EquivalentSource")
            if not nid or eq is None:
                continue
            nodes.add(nid)
            kvll_txt = eq.findtext("KVLL")
            if kvll_txt not in (None, ""):
                try:
                    kvll_map[nid] = float(kvll_txt) * 1000.0  # kV -> V
                except Exception:
                    pass
    return nodes, kvll_map


def _parse_bus_rows(input_path: Path) -> List[Dict]:
    """
    Build rows for the Bus sheet.

    Columns:
      Bus | Base Voltage (V) | Initial Vmag | Unit | Angle | Type

    Rules:
      - Buses with no ACTIVE connections are prefixed with '//' (commented out).
        ACTIVE = branch with both endpoints that will exist on the Bus sheet.
      - Only sources that appear on the Voltage Source page (Topo Substation & not EquivalentMode=1)
        are SLACK. Feeder-start pseudo-sources that don't appear on that page remain PQ.
      - Unused buses are commented even if they are SLACK.
      - All buses reachable from a VS-page source *without crossing a transformer* get LN = KVLL/√3
        for that source; others default to 7.2 kV LN.
      - Island policy:
          * If an Active Island is chosen → keep only that island (even if it has no source).
          * If none chosen → keep only islands with a voltage source.
    """
    root = _read_xml(Path(input_path))

    # --- Only consider sources that appear on the Voltage Source page ---
    vs_slack_nodes, kvll_map = _gather_vs_page_sources_and_kvll(root)

    # Standard phase angles for display
    phase_ang = {"A": 0.0, "B": -120.0, "C": 120.0}

    # -------- PASS 1: scan all sections to determine endpoints/usage and detect local pseudos
    BRANCH_TAGS = [
        # Overhead / Underground lines (all flavors)
        "OverheadLine", "OverheadLineUnbalanced", "OverheadByPhase",
        "Underground", "UndergroundCable", "UndergroundCableUnbalanced", "UndergroundByPhase",
        # Transformers
        "Transformer",
        # Switching / protection devices (treat like branches)
        "Switch", "Fuse", "Recloser", "Breaker", "Sectionalizer", "Isolator",
    ]
    # For adjacency used to propagate HV from source to transformer, exclude transformers:
    BRANCH_NO_XFMR_TAGS = [
        "OverheadLine", "OverheadLineUnbalanced", "OverheadByPhase",
        "Underground", "UndergroundCable", "UndergroundCableUnbalanced", "UndergroundByPhase",
        "Switch", "Fuse", "Recloser", "Breaker", "Sectionalizer", "Isolator",
    ]

    branch_endpoints: Set[str] = set()           # all From/To seen in branch sections (sanitized)
    all_from_nodes: Set[str] = set()             # all FromNodeID across sections (sanitized)
    device_terminal_candidates: Set[str] = set() # ToNodeID of SpotLoad/Shunt sections (sanitized)

    # For branch-local pseudo detection:
    from_count: Dict[str, int] = {}
    to_count_nonlocal: Dict[str, int] = {}       # times a node is To in a section where it's NOT local pseudo
    local_pseudo_to_candidates: Set[str] = set() # ToNodeIDs equal to local SectionID/DeviceNumber/DeviceID (sanitized)

    # We still gather a "raw" degree during scan, but final "active" usage is computed later.
    raw_degree: Dict[str, int] = {}

    # Also build an adjacency (no-transformer edges) for HV propagation
    adj: Dict[str, Set[str]] = {}

    def _add_edge(a: str, b: str) -> None:
        if not a or not b:
            return
        adj.setdefault(a, set()).add(b)
        adj.setdefault(b, set()).add(a)

    for sec in root.findall(".//Sections/Section"):
        f_raw = (sec.findtext("./FromNodeID") or "").strip()
        t_raw = (sec.findtext("./ToNodeID") or "").strip()
        f = safe_name(f_raw)
        t = safe_name(t_raw)

        if f:
            all_from_nodes.add(f)
            from_count[f] = from_count.get(f, 0) + 1

        has_branch = _has_any(sec, BRANCH_TAGS)
        has_branch_no_xfmr = _has_any(sec, BRANCH_NO_XFMR_TAGS)
        has_spot = sec.find(".//Devices/SpotLoad") is not None
        has_shunt = (
            sec.find(".//Devices/ShuntCapacitor") is not None
            or sec.find(".//Devices/ShuntReactor") is not None
        )

        if has_branch:
            if f:
                branch_endpoints.add(f)
                raw_degree[f] = raw_degree.get(f, 0) + 1
            if t:
                branch_endpoints.add(t)
                raw_degree[t] = raw_degree.get(t, 0) + 1
            # local pseudo test for the To side of this branch
            if t:
                lp = _local_pseudos(sec)  # already sanitized inside
                if t in lp:
                    local_pseudo_to_candidates.add(t)
                else:
                    to_count_nonlocal[t] = to_count_nonlocal.get(t, 0) + 1

        if has_spot or has_shunt:
            # treat the network side (FromNodeID) as a usage (bump raw_degree)
            if f:
                raw_degree[f] = raw_degree.get(f, 0) + 1
            # ToNodeID on device sections often refers to a local pseudo terminal — don't bump degree on t
            if t:
                device_terminal_candidates.add(t)

        # Build adjacency only for *non-transformer* branches
        if has_branch_no_xfmr and f and t:
            _add_edge(f, t)

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
        nid = safe_name(nid)
        if not nid:
            return False
        if nid in terminal_exclusions:
            return False
        return True

    node_phases: Dict[str, Set[str]] = {}

    def _add(nid: str | None, pstr: str) -> None:
        nid = safe_name(nid)
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

        # Fallback: conservative (may create buses with no active usage)
        _add(f, ph)

    # -------- PASS 3: compute ACTIVE usage
    known_bases: Set[str] = set(node_phases.keys())
    active_degree: Dict[str, int] = {}

    for sec in root.findall(".//Sections/Section"):
        f = safe_name(sec.findtext("./FromNodeID"))
        t = safe_name(sec.findtext("./ToNodeID"))

        has_branch = _has_any(sec, BRANCH_TAGS)
        has_spot   = sec.find(".//Devices/SpotLoad") is not None
        has_shunt  = (
            sec.find(".//Devices/ShuntCapacitor") is not None
            or sec.find(".//Devices/ShuntReactor") is not None
        )

        if has_branch:
            # Only count as ACTIVE if BOTH endpoints will exist on Bus sheet
            if f and t and (f in known_bases) and (t in known_bases):
                active_degree[f] = active_degree.get(f, 0) + 1
                active_degree[t] = active_degree.get(t, 0) + 1

        if has_spot or has_shunt:
            # Count the network side only if that bus will exist
            if f and (f in known_bases):
                active_degree[f] = active_degree.get(f, 0) + 1

    # -------- Build HV propagation from VS-page sources up to transformer (no-XFMR graph)
    # Map: node -> ln_volts assigned from a particular source's KVLL/√3
    hv_ln_assign: Dict[str, float] = {}

    from collections import deque

    for src in vs_slack_nodes:
        if src not in kvll_map:
            continue
        ln_volts = kvll_map[src] / (3 ** 0.5)
        # BFS from this source over the no-transformer adjacency
        q = deque()
        visited: Set[str] = set()
        if src in adj:
            q.append(src)
            visited.add(src)
        else:
            # Even if no neighbors, still tag the source itself if present.
            visited.add(src)
        while q:
            u = q.popleft()
            # Assign if bus exists on the Bus sheet
            if u in known_bases and u not in hv_ln_assign:
                hv_ln_assign[u] = ln_volts
            for v in adj.get(u, ()):
                if v not in visited:
                    visited.add(v)
                    q.append(v)
        # Also tag the source itself even if it didn’t appear in adj
        if src in known_bases and src not in hv_ln_assign:
            hv_ln_assign[src] = ln_volts

    # -------- Emit rows (comment out unused OR island-filtered buses; SLACK only for VS-page sources)
    def pkey(p: str) -> int:
        return PHASES.index(p)

    rows: List[Dict] = []
    for node in sorted(node_phases):
        is_active = active_degree.get(node, 0) > 0

        # Only sources that appear on the Voltage Source page are SLACK
        bus_type = "SLACK" if node in vs_slack_nodes else "PQ"

        # Island policy decides if this node should be commented (unless it's unused, which also comments)
        island_comment = should_comment_bus(node)

        for ph in sorted(node_phases[node], key=pkey):
            bus_name = f"{node}_{ph.lower()}"

            # Comment if unused or excluded by island policy
            if (not is_active) or island_comment:
                bus_name = "//" + bus_name

            # Per-row voltages:
            # - If node is reachable from a VS-page source without crossing a transformer, use that source's LN
            # - Else default to 7.2 kV LN
            v_ln = hv_ln_assign.get(node, 7200.0)

            rows.append(
                {
                    "Bus": bus_name,
                    "Base Voltage (V)": v_ln,   # per-phase base
                    "Initial Vmag": v_ln,       # equal to base (per phase)
                    "Unit": "V",
                    "Angle": phase_ang.get(ph, 0.0),
                    "Type": bus_type,
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
