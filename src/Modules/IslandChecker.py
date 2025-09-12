# Modules/IslandChecker.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Set, Tuple
import xml.etree.ElementTree as ET

from Modules.General import safe_name, set_island_context

# Device groups considered as *topology* edges between FromNodeID <-> ToNodeID
LINE_LIKE = {
    "OverheadLine", "OverheadLineUnbalanced", "OverheadByPhase",
    "Underground", "UndergroundCable", "UndergroundCableUnbalanced", "UndergroundByPhase",
    "Cable",  # include standalone Cable device as conducting branch
}
SWITCH_LIKE = {"Switch", "Sectionalizer", "Breaker", "Fuse", "Recloser", "Isolator", "Miscellaneous"}
TRANSFORMERS = {"Transformer"}


def _read_xml(path: Path) -> ET.Element:
    return ET.fromstring(Path(path).read_text(encoding="utf-8", errors="ignore"))


def _dev_is_closed(dev: ET.Element) -> bool:
    """
    Heuristic for 'closed/in-service':
      - ConnectionStatus == 'Disconnected' -> OPEN
      - NormalStatus == 'open'            -> OPEN
      - ClosedPhase present and not 'None' -> CLOSED
      - Otherwise default to CLOSED
    """
    cs = (dev.findtext("ConnectionStatus") or "").strip().lower()
    if cs == "disconnected":
        return False

    ns = (dev.findtext("NormalStatus") or "").strip().lower()
    if ns == "open":
        return False

    cp = (dev.findtext("ClosedPhase") or "").strip().upper()
    if cp and cp not in ("", "NONE"):
        return True

    return True


def _section_has_closed_connection(sec: ET.Element) -> bool:
    """
    Decide if the section provides a closed conducting path between From/To.

    Important nuance for CYME sections that include both a line and a switch:
    - An OPEN switch in the section should OPEN the whole section (no connection),
      even if a line-like device is 'in service'. In practice, devices in a
      section are in series for topology purposes.
    - If there is a switch-only tie section, a CLOSED switch should connect.

    Therefore:
      - If ANY switch-like device is present and OPEN -> section is OPEN.
      - Else, if ANY device (switch/line/transformer) is CLOSED -> section is CLOSED.
      - Else -> OPEN.
    """
    devs = sec.find("./Devices")
    if devs is None:
        return False

    # If any switch-like device is OPEN, the section is open
    for tag in SWITCH_LIKE:
        for d in devs.findall(tag):
            if not _dev_is_closed(d):
                return False

    # Otherwise, if any device that can conduct is CLOSED, the section is closed
    for tag in TRANSFORMERS:
        for d in devs.findall(tag):
            if _dev_is_closed(d):
                return True

    for tag in LINE_LIKE:
        for d in devs.findall(tag):
            if _dev_is_closed(d):
                return True

    for tag in SWITCH_LIKE:
        for d in devs.findall(tag):
            if _dev_is_closed(d):
                return True

    return False


def _vs_page_source_nodes(root: ET.Element) -> Set[str]:
    """
    Return the set of buses that appear as Voltage Sources on the VS page:
    - Only Topo blocks where NetworkType == 'Substation'
    - EquivalentMode != '1'
    - And the Source has an EquivalentSource model (same criteria Bus/Voltage_Source use)
    """
    out: Set[str] = set()
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
            # Must have an EquivalentSource block to show on VS page
            eq = src.find("./EquivalentSourceModels/EquivalentSourceModel/EquivalentSource")
            if nid and eq is not None:
                out.add(nid)
    # Fallback: if no Topo/Substation sources were found, keep behavior aligned
    # with the VS sheet and return empty set rather than all .//Sources/Source
    return out


def _shunt_buses(root: ET.Element) -> Set[str]:
    out: Set[str] = set()
    for sec in root.findall(".//Sections/Section"):
        devs = sec.find("./Devices")
        if devs is None:
            continue
        if devs.find("ShuntCapacitor") is not None or devs.find("ShuntReactor") is not None:
            fb = safe_name(sec.findtext("FromNodeID"))
            if fb:
                out.add(fb)
    return out


def _build_graph(root: ET.Element) -> Tuple[Dict[str, Set[str]], int, int]:
    """Undirected graph of sanitized bus names using closed devices only."""
    adj: Dict[str, Set[str]] = {}
    edges_closed = 0
    edges_open_ignored = 0

    for sec in root.findall(".//Sections/Section"):
        fb = safe_name(sec.findtext("FromNodeID"))
        tb = safe_name(sec.findtext("ToNodeID"))
        if not fb or not tb:
            continue

        if _section_has_closed_connection(sec):
            adj.setdefault(fb, set()).add(tb)
            adj.setdefault(tb, set()).add(fb)
            edges_closed += 1
        else:
            edges_open_ignored += 1

        # Ensure standalone nodes exist
        adj.setdefault(fb, adj.get(fb, set()))
        adj.setdefault(tb, adj.get(tb, set()))

    return adj, edges_closed, edges_open_ignored


def _components(adj: Dict[str, Set[str]]) -> List[Set[str]]:
    """Connected components via DFS."""
    seen: Set[str] = set()
    comps: List[Set[str]] = []
    for v in adj:
        if v in seen:
            continue
        stack = [v]
        comp: Set[str] = set()
        while stack:
            u = stack.pop()
            if u in seen:
                continue
            seen.add(u)
            comp.add(u)
            stack.extend(w for w in adj[u] if w not in seen)
        comps.append(comp)
    return comps


def check_islands(xml_path: Path) -> Dict:
    """
    Returns:
      {
        'count': int,
        'components': [
           {'index': i, 'size': n, 'nodes': [...], 'limited_node_sample': [...],
            'has_source': bool, 'has_shunt': bool}
        ],
        'edges_closed': int,
        'edges_open_ignored': int,
        'nodes_total': int
      }
    """
    root = _read_xml(xml_path)
    adj, e_closed, e_ignored = _build_graph(root)
    comps = _components(adj)

    source_nodes = _vs_page_source_nodes(root)
    shunt_nodes = _shunt_buses(root)

    # Sort "good" islands (with sources) first, then by size desc, then lexicographically
    comps_info = []
    for comp in comps:
        has_source = any(n in source_nodes for n in comp)
        comps_info.append((comp, has_source))
    comps_info.sort(key=lambda t: (0 if t[1] else 1, -len(t[0]), min(t[0]) if t[0] else ""))

    out_list = []
    for i, (comp, has_source) in enumerate(comps_info, start=1):
        has_shunt = any(n in shunt_nodes for n in comp)
        sample = sorted(list(comp))[:20]
        out_list.append({
            "index": i,
            "size": len(comp),
            "nodes": sorted(list(comp)),
            "limited_node_sample": sample,
            "has_source": has_source,
            "has_shunt": has_shunt,
        })

    return {
        "count": len(comps),
        "components": out_list,
        "edges_closed": e_closed,
        "edges_open_ignored": e_ignored,
        "nodes_total": len(adj),
    }


def log_islands(xml_path: Path, per_island_limit: int | None = None) -> None:
    """Console-friendly vertical printout."""
    s = check_islands(xml_path)
    print(f"[Islands] Count={s['count']}  Nodes={s['nodes_total']}  "
          f"ClosedEdges={s['edges_closed']}  OpenIgnored={s['edges_open_ignored']}")
    print("-" * 72)
    for comp in s["components"]:
        src = "Yes" if comp["has_source"] else "No"
        sh  = "Yes" if comp["has_shunt"]  else "No"
        print(f"Island {comp['index']}  |  Size: {comp['size']}  |  Source: {src}  |  Shunt: {sh}")
        print("  nodes:")
        nodes = comp["nodes"]
        if per_island_limit is not None and len(nodes) > per_island_limit:
            for n in nodes[:per_island_limit]:
                print(f"    - {n}")
            print(f"    ... (+{len(nodes) - per_island_limit} more)")
        else:
            for n in nodes:
                print(f"    - {n}")
        print("")


def build_island_context(xml_path: Path) -> dict:
    """
    Build a context writers can use:
      {
        'bus_to_island': {bus_base: island_idx, ...},
        'bad_buses': set(bus_base, ...),            # reserved for truly bad pseudo terminals (left empty by default)
        'slack_per_island': {island_idx: bus_base}, # exactly one per island WITH a source
        'islands': {island_idx: set(bus_base, ...)},
        'sourceful_islands': set([island_idx, ...]) # convenience for UI/filters
      }
    """
    s = check_islands(xml_path)

    # Map each bus -> island index; collect island sets
    bus_to_island: Dict[str, int] = {}
    islands: Dict[int, Set[str]] = {}
    has_source_by_island: Dict[int, bool] = {}
    for comp in s["components"]:
        idx = comp["index"]
        bases = set(comp["nodes"])
        islands[idx] = bases
        for b in bases:
            bus_to_island[b] = idx
        has_source_by_island[idx] = bool(comp["has_source"])

    # Source nodes used for slack selection (VS page sources only)
    root = _read_xml(xml_path)
    source_nodes = _vs_page_source_nodes(root)

    slack_per_island: Dict[int, str] = {}
    for idx, bases in islands.items():
        if not has_source_by_island.get(idx, False):
            continue  # only assign slack if the island really has a source
        prefer = sorted(bases & source_nodes)
        slack_per_island[idx] = prefer[0] if prefer else sorted(bases)[0]

    # IMPORTANT CHANGE:
    # Do NOT pre-mark all buses in no-source islands as "bad".
    # Leave this set for intrinsic pseudo/terminal nodes only (empty by default).
    bad_buses: Set[str] = set()

    sourceful_islands = {i for i, ok in has_source_by_island.items() if ok}

    return {
        "bus_to_island": bus_to_island,
        "bad_buses": bad_buses,
        "slack_per_island": slack_per_island,
        "islands": islands,
        "sourceful_islands": sourceful_islands,
    }


def analyze_and_set_island_context(xml_path: Path, *, per_island_limit: int | None = None) -> dict:
    """Print vertical summary and store context globally for writers."""
    log_islands(xml_path, per_island_limit=per_island_limit)
    ctx = build_island_context(xml_path)
    set_island_context(ctx)
    return ctx
