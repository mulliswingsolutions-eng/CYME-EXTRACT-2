# Modules/Bus.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Set, Tuple
import xml.etree.ElementTree as ET

# replace with
from Modules.General import safe_name, get_island_context
# (optional robust fallback for standalone use)
try:
    from Modules.General import get_island_context  # already imported above; keeps Pylance happy
except Exception:
    def get_island_context() -> dict:
        return {}
    
PHASES = ("A", "B", "C")

# Add just once in Bus.py
def _fnum(x: str | None) -> float | None:
    try:
        if x is None:
            return None
        s = x.strip()
        if not s:
            return None
        if s.lower().endswith("deg"):  # tolerate "330deg"
            s = s[:-3]
        return float(s)
    except Exception:
        return None


def _collect_transformers_with_kvll(root: ET.Element) -> list[tuple[str, str, float | None, float | None]]:
    """
    Returns list of (from_bus, to_bus, KVLL_primary_volts, KVLL_secondary_volts).
    Uses TransformerDB when available; falls back to section fields if present.
    """
    # Map EquipmentID/DeviceID -> (kvp_ll, kvs_ll) in volts (LL)
    db: dict[str, tuple[float | None, float | None]] = {}
    for tdb in root.findall(".//TransformerDB"):
        eid = safe_name(tdb.findtext("EquipmentID"))
        kvp = _fnum(tdb.findtext("PrimaryVoltage")) or _fnum(tdb.findtext("PrimaryKV"))
        kvs = _fnum(tdb.findtext("SecondaryVoltage")) or _fnum(tdb.findtext("SecondaryKV"))
        if kvp is not None:
            kvp *= 1000.0
        if kvs is not None:
            kvs *= 1000.0
        if eid:
            db[eid] = (kvp, kvs)

    pairs: list[tuple[str, str, float | None, float | None]] = []
    for sec in root.findall(".//Sections/Section"):
        xf = sec.find(".//Devices/Transformer")
        if xf is None:
            continue
        fb = safe_name(sec.findtext("./FromNodeID"))
        tb = safe_name(sec.findtext("./ToNodeID"))
        if not fb or not tb:
            continue

        dev_id = safe_name(xf.findtext("DeviceID"))
        kvp_ll = kvs_ll = None

        # Prefer DB by DeviceID/EquipmentID
        if dev_id and dev_id in db:
            kvp_ll, kvs_ll = db[dev_id]

        # Fallbacks from the section if DB was missing or 0
        if kvp_ll is None:
            kvp_ll = _fnum(xf.findtext("SystemBaseVoltage/Primary")) or _fnum(xf.findtext("PrimaryKV"))
            if kvp_ll is not None:
                kvp_ll *= 1000.0
        if kvs_ll is None:
            kvs_ll = _fnum(xf.findtext("SystemBaseVoltage/Secondary")) or _fnum(xf.findtext("SecondaryKV"))
            if kvs_ll is not None:
                kvs_ll *= 1000.0

        pairs.append((fb, tb, kvp_ll, kvs_ll))
    return pairs


def _propagate_ln_via_transformers(
    hv_ln_assign: dict[str, float],
    adj_no_xfmr: dict[str, set[str]],
    xf_pairs: list[tuple[str, str, float | None, float | None]],
) -> None:
    """
    Fill in LN volts on the far side of transformers, then flood across the
    non-transformer graph. Repeat until no changes (handles cascaded XFRs).
    """
    import math
    def close(a: float, b: float, rel: float = 0.06) -> bool:
        m = max(abs(a), abs(b), 1.0)
        return abs(a - b) / m <= rel

    def flood(start_node: str, ln_val: float) -> None:
        # Push the assigned LN across NO-TRANSFORMER edges
        stack = [start_node]
        while stack:
            u = stack.pop()
            for v in adj_no_xfmr.get(u, ()):
                if v not in hv_ln_assign:
                    hv_ln_assign[v] = ln_val
                    stack.append(v)

    changed = True
    while changed:
        changed = False
        for fb, tb, kvp_ll, kvs_ll in xf_pairs:
            lf = hv_ln_assign.get(fb)
            lt = hv_ln_assign.get(tb)
            # Need at least one side assigned and both KVLLs to decide orientation
            if (lf is None and lt is None) or (kvp_ll is None or kvs_ll is None):
                continue

            kvp_ln = kvp_ll / math.sqrt(3.0) if kvp_ll is not None else None
            kvs_ln = kvs_ll / math.sqrt(3.0) if kvs_ll is not None else None
            if kvp_ln is None or kvs_ln is None:
                continue

            # If From side looks like primary, set To as secondary (and propagate)
            if lf is not None and lt is None and (close(lf, kvp_ln) or not close(lf, kvs_ln)):
                hv_ln_assign[tb] = kvs_ln
                flood(tb, kvs_ln)
                changed = True
                continue

            # If From side looks like secondary, set To as primary
            if lf is not None and lt is None and close(lf, kvs_ln):
                hv_ln_assign[tb] = kvp_ln
                flood(tb, kvp_ln)
                changed = True
                continue

            # Mirror: To side known, From side unknown
            if lt is not None and lf is None and (close(lt, kvp_ln) or not close(lt, kvs_ln)):
                hv_ln_assign[fb] = kvs_ln
                flood(fb, kvs_ln)
                changed = True
                continue

            if lt is not None and lf is None and close(lt, kvs_ln):
                hv_ln_assign[fb] = kvp_ln
                flood(fb, kvp_ln)
                changed = True
                continue

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

def _gather_vs_page_sources_and_kvll(root: ET.Element) -> tuple[set[str], dict[str, float]]:
    """
    Return ({source_nodes}, {node -> KVLL_volts}) for sources that WILL SHOW
    on the Voltage Source page:
      - Topo where NetworkType == 'Substation'
      - EquivalentMode != '1'
      - Source must have an EquivalentSource block
    """
    nodes: set[str] = set()
    kvll_map: dict[str, float] = {}
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
      - Only sources that appear on the Voltage Source page (Topo Substation & not EquivalentMode=1)
        are marked SLACK.
      - Islands without sources are commented via bad_buses in the island context.
      - Base LN voltage is propagated from each VS-page source through the network
        WITHOUT crossing transformers; i.e., KV only changes at a transformer.
    """
    root = _read_xml(Path(input_path))

    # angles for A/B/C (display only)
    phase_ang = {"A": 0.0, "B": -120.0, "C": 120.0}

    # -------- source set & kVLL (from Voltage Source page only)
    vs_slack_nodes, kvll_map = _gather_vs_page_sources_and_kvll(root)

    # -------- PASS 1: scan sections (same pseudo filtering you had)
    BRANCH_TAGS = [
        "OverheadLine", "OverheadLineUnbalanced", "OverheadByPhase",
        "Underground", "UndergroundCable", "UndergroundCableUnbalanced", "UndergroundByPhase",
        "Transformer",
        "Switch", "Fuse", "Recloser", "Breaker", "Sectionalizer", "Isolator",
        "Miscellaneous",  # NEW: allow endpoints/degree counting across meter/LA, etc.
    ]
    # Edges for voltage propagation: **exclude transformers**
    BRANCH_NO_XFMR_TAGS = [
        "OverheadLine", "OverheadLineUnbalanced", "OverheadByPhase",
        "Underground", "UndergroundCable", "UndergroundCableUnbalanced", "UndergroundByPhase",
        "Switch", "Fuse", "Recloser", "Breaker", "Sectionalizer", "Isolator",
        "Miscellaneous",  # NEW: treat as a conducting link (no voltage change)
    ]

    branch_endpoints: Set[str] = set()
    all_from_nodes: Set[str] = set()
    device_terminal_candidates: Set[str] = set()

    from_count: Dict[str, int] = {}
    to_count_nonlocal: Dict[str, int] = {}
    local_pseudo_to_candidates: Set[str] = set()

    raw_degree: Dict[str, int] = {}   # <-- ADD THIS LINE
    xf_endpoints: Set[str] = set()

    # adjacency for BFS (non-transformer branches only)
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

        if sec.find(".//Devices/Transformer") is not None:
            if f:
                xf_endpoints.add(f)
            if t:
                xf_endpoints.add(t)

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
            if t:
                lp = _local_pseudos(sec)
                if t in lp:
                    local_pseudo_to_candidates.add(t)
                else:
                    to_count_nonlocal[t] = to_count_nonlocal.get(t, 0) + 1

        if has_spot or has_shunt:
            if f:
                raw_degree[f] = raw_degree.get(f, 0) + 1
            if t:
                device_terminal_candidates.add(t)

        # Build **non-transformer** adjacency
        if has_branch_no_xfmr and f and t:
            _add_edge(f, t)

    terminal_exclusions_1: Set[str] = {
        n for n in device_terminal_candidates if n not in branch_endpoints and n not in all_from_nodes
    }
    terminal_exclusions_2: Set[str] = {
        n for n in local_pseudo_to_candidates if from_count.get(n, 0) == 0 and to_count_nonlocal.get(n, 0) == 0
    }
    terminal_exclusions: Set[str] = (terminal_exclusions_1 | terminal_exclusions_2) - xf_endpoints

    # -------- PASS 2: node -> phases map (with exclusions)
    def _is_real_bus(nid: str | None) -> bool:
        nid = safe_name(nid)
        if not nid:
            return False
        if nid in xf_endpoints:  # always keep transformer ends
            return True
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
            _add(f, ph)
            continue

        if has_branch:
            _add(f, ph)
            _add(t, ph)
            continue

        _add(f, ph)

    known_bases: Set[str] = set(node_phases.keys())

    # -------- PASS 3: ACTIVE usage (for commenting rules)
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
            if f and t and (f in known_bases) and (t in known_bases):
                active_degree[f] = active_degree.get(f, 0) + 1
                active_degree[t] = active_degree.get(t, 0) + 1

        if has_spot or has_shunt:
            if f and (f in known_bases):
                active_degree[f] = active_degree.get(f, 0) + 1

    # -------- LN voltage assignment by BFS from VS-page sources (no transformers)
    hv_ln_assign: Dict[str, float] = {}
    from collections import deque

    for src in vs_slack_nodes:
        kvll_v = kvll_map.get(src)
        if kvll_v is None:
            continue
        ln_v = kvll_v / (3 ** 0.5)
        q = deque()
        visited: Set[str] = set()
        if src in adj:
            q.append(src)
            visited.add(src)
        else:
            visited.add(src)
        while q:
            u = q.popleft()
            if u in known_bases and u not in hv_ln_assign:
                hv_ln_assign[u] = ln_v
            for v in adj.get(u, ()):
                if v not in visited:
                    visited.add(v)
                    q.append(v)
        # tag the source itself even if isolated
        if src in known_bases and src not in hv_ln_assign:
            hv_ln_assign[src] = ln_v
            
    xf_pairs = _collect_transformers_with_kvll(root)
    _propagate_ln_via_transformers(hv_ln_assign, adj, xf_pairs)

    # -------- Island context
    ctx = get_island_context() or {}
    bad_buses: Set[str] = set(ctx.get("bad_buses", set()))
    bus_to_island: Dict[str, int] = dict(ctx.get("bus_to_island", {}))
    slack_per_island: Dict[int, str] = dict(ctx.get("slack_per_island", {}))

    # -------- Emit rows
    def pkey(p: str) -> int:
        return PHASES.index(p)

    rows: List[Dict] = []
    for node in sorted(node_phases):
        island = bus_to_island.get(node)
        is_in_bad_island = node in bad_buses
        is_active = active_degree.get(node, 0) > 0

        # SLACK only if this node is an actual VS-page source (or island slack selected in context)
        if node in vs_slack_nodes:
            bus_type = "SLACK"
        elif island is not None and slack_per_island.get(island) == node and not is_in_bad_island:
            bus_type = "SLACK"
        else:
            bus_type = "PQ"

        # per-node base LN volts: use BFS assignment; if not assigned, default (e.g., 7.2 kV LN)
        node_ln_v = hv_ln_assign.get(node, 7200.0)

        for ph in sorted(node_phases[node], key=pkey):
            bus_name = f"{node}_{ph.lower()}"

            # comment unused or bad-island buses
            should_comment = is_in_bad_island or (bus_type != "SLACK" and not is_active)
            if should_comment:
                bus_name = "//" + bus_name

            rows.append(
                {
                    "Bus": bus_name,
                    "Base Voltage (V)": node_ln_v,   # per-phase base LN volts
                    "Initial Vmag": node_ln_v,       # equal to base for initialization
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
