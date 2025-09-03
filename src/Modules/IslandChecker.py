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
    devs = sec.find("./Devices")
    if devs is None:
        return False

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


def _sources_nodes(root: ET.Element) -> Set[str]:
    out: Set[str] = set()
    for src in root.findall(".//Sources/Source"):
        nid = safe_name(src.findtext("SourceNodeID"))
        if nid:
            out.add(nid)
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

    source_nodes = _sources_nodes(root)
    shunt_nodes = _shunt_buses(root)

    out_list = []
    for i, comp in enumerate(sorted(comps, key=lambda s: (-len(s), min(s) if s else "")), start=1):
        has_source = any(n in source_nodes for n in comp)
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

    # Source nodes (prefer slack there)
    root = _read_xml(xml_path)
    source_nodes = set()
    for src in root.findall(".//Sources/Source"):
        nid = safe_name(src.findtext("SourceNodeID"))
        if nid:
            source_nodes.add(nid)

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
