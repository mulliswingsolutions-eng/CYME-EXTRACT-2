# Modules/Load.py
from __future__ import annotations

from pathlib import Path
import math
import re
import xml.etree.ElementTree as ET
from collections import deque
from typing import Dict, List, Tuple, Any, Set
from Modules.General import safe_name
from Modules.Bus import extract_bus_data   # <-- NEW: to get active Bus bases

PHASES = ("A", "B", "C")
PHASE_SUFFIX = {"A": "_a", "B": "_b", "C": "_c"}
EPS = 1e-6  # numerical zero

_PHASE_SUFFIX_RE = re.compile(r"_(a|b|c)$")

def _read_xml(path: Path) -> ET.Element:
    return ET.fromstring(path.read_text(encoding="utf-8", errors="ignore"))


# =======================
# Voltage map (KVLL) builder
# =======================
def _get_source_kvll(root: ET.Element) -> float:
    eq = root.find(".//Sources/Source/EquivalentSourceModels/EquivalentSourceModel/EquivalentSource")
    if eq is not None:
        try:
            v = eq.findtext("KVLL")
            return float(v) if v not in (None, "") else 0.0
        except Exception:
            return 0.0
    return 0.0


def _read_transformer_db(root: ET.Element) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for tdb in root.findall(".//TransformerDB"):
        eid = (tdb.findtext("EquipmentID") or "").strip()
        if not eid:
            continue

        def f(name: str) -> float:
            v = tdb.findtext(name)
            try:
                return float(v) if v not in (None, "") else 0.0
            except Exception:
                return 0.0

        out[eid] = {
            "kvp": f("PrimaryVoltage") or f("PrimaryKV"),
            "kvs": f("SecondaryVoltage") or f("SecondaryKV"),
        }
    return out


def _build_voltage_map(txt_path: Path) -> Dict[str, float]:
    """
    Seed KVLL from <Sources> and transformer DB, then propagate through
    'pure network' branches (lines/cables/switches) that do not host loads/shunts.
    All node IDs are sanitized so they match sheet outputs.
    """
    root = _read_xml(Path(txt_path))
    kvll_default = _get_source_kvll(root)
    tdb = _read_transformer_db(root)

    G: Dict[str, Set[str]] = {}

    def add_edge(a_raw: str, b_raw: str) -> None:
        a, b = safe_name(a_raw), safe_name(b_raw)
        if not a or not b:
            return
        G.setdefault(a, set()).add(b)
        G.setdefault(b, set()).add(a)

    seeds: Dict[str, float] = {}
    src_node = safe_name(root.findtext(".//Sources/Source/SourceNodeID"))
    if src_node and kvll_default > 0:
        seeds[src_node] = kvll_default

    for sec in root.findall(".//Sections/Section"):
        f_raw = (sec.findtext("./FromNodeID") or "").strip()
        t_raw = (sec.findtext("./ToNodeID") or "").strip()
        f = safe_name(f_raw)
        t = safe_name(t_raw)

        has_spot = sec.find(".//Devices/SpotLoad") is not None
        has_dist = sec.find(".//Devices/DistributedLoad") is not None   # NEW: distributed
        has_shunt = (sec.find(".//Devices/ShuntCapacitor") is not None
                     or sec.find(".//Devices/ShuntReactor") is not None)

        xf = sec.find(".//Devices/Transformer")
        if xf is not None:
            dev_id = (xf.findtext("DeviceID") or "").strip()
            normal = safe_name(xf.findtext("NormalFeedingNodeID"))
            vals = tdb.get(dev_id, {})
            kvp = vals.get("kvp") or 0.0
            kvs = vals.get("kvs") or 0.0
            if kvp > 0 or kvs > 0:
                if normal and normal == f:
                    if kvp > 0: seeds.setdefault(f, kvp)
                    if kvs > 0: seeds.setdefault(t, kvs)
                elif normal and normal == t:
                    if kvp > 0: seeds.setdefault(t, kvp)
                    if kvs > 0: seeds.setdefault(f, kvs)
                else:
                    if kvp > 0: seeds.setdefault(f, kvp)
                    if kvs > 0: seeds.setdefault(t, kvs)
            continue

        has_line = (
            sec.find(".//Devices/OverheadLine") is not None
            or sec.find(".//Devices/OverheadLineUnbalanced") is not None
            or sec.find(".//Devices/OverheadByPhase") is not None
            or sec.find(".//Devices/UndergroundCable") is not None
            or sec.find(".//Devices/Underground") is not None
        )
        has_switch = sec.find(".//Devices/Switch") is not None

        # Do NOT connect across sections that host any load (spot or distributed) or shunt
        if (has_line or has_switch) and not (has_spot or has_dist or has_shunt):
            add_edge(f, t)

    # BFS propagate (on sanitized node names)
    bus_kv: Dict[str, float] = {}

    for start, kv in seeds.items():
        if start in bus_kv:
            continue
        q = deque([start])
        bus_kv[start] = kv
        while q:
            u = q.popleft()
            for v in G.get(u, ()):
                if v not in bus_kv:
                    bus_kv[v] = kv
                    q.append(v)

    bus_kv["_default_"] = kvll_default if kvll_default > 0 else 0.0
    return bus_kv


# =======================
# Load parsing
# =======================
def _sanitize_id(s: str) -> str:
    return safe_name(s)


def _norm_load_id(device_number: str, section_id: str, from_node_id: str) -> str:
    for src in (device_number, section_id, from_node_id):
        src = (src or "").strip()
        if src:
            return f"LD_{_sanitize_id(src)}"
    return "LD_unknown"


def _get_type_attr(elem: ET.Element) -> str:
    t = elem.get("Type")
    if not t:
        for k, v in elem.attrib.items():
            if k.endswith("}type"):
                t = v
                break
    return (t or "").upper()


def _float_or_none(x: str | None) -> float | None:
    try:
        return float(x) if x not in (None, "") else None
    except Exception:
        return None


def _kw_kvar_from_value(val: ET.Element) -> tuple[float | None, float | None]:
    """
    Returns (kW, kVAr) using the best-available fields.
    Supports:
      - KW_KVAR
      - KW_PF  (fallback to ConnectedKVA*PF when kW ~ 0 but KVA present)
      - KVA_PF
      - KVA_KVAR
    """
    lv = val.find("./LoadValue")
    if lv is None:
        return None, None

    t = _get_type_attr(lv)
    conn_kva = _float_or_none(val.findtext("ConnectedKVA"))

    if t.endswith("KW_KVAR"):
        return _float_or_none(lv.findtext("KW")), _float_or_none(lv.findtext("KVAR"))

    if t.endswith("KW_PF"):
        kw = _float_or_none(lv.findtext("KW"))
        pf_raw = _float_or_none(lv.findtext("PF"))
        if pf_raw is None or pf_raw <= 0:
            return kw, None
        pf = pf_raw / 100.0 if pf_raw > 1.0 else pf_raw
        pf = max(0.0, min(1.0, pf))
        # normal case: derive kvar from kw & pf
        if kw is not None and abs(kw) > EPS:
            try:
                phi = math.acos(pf)
                kvar = kw * math.tan(phi)
                return kw, kvar
            except Exception:
                return kw, None
        # fallback: use ConnectedKVA if given (handles files with KW=0 but KVA>0)
        if conn_kva is not None and conn_kva > EPS:
            p = conn_kva * pf
            try:
                phi = math.acos(pf)
                q = conn_kva * math.sin(phi)
            except Exception:
                q = None
            return p, q
        return kw, None

    if t.endswith("KVA_PF"):
        kva = _float_or_none(lv.findtext("KVA"))
        pf_raw = _float_or_none(lv.findtext("PF"))
        if kva is None or pf_raw is None or pf_raw <= 0:
            return None, None
        pf = pf_raw / 100.0 if pf_raw > 1.0 else pf_raw
        pf = max(0.0, min(1.0, pf))
        p = kva * pf
        try:
            phi = math.acos(pf)
            q = kva * math.sin(phi)
        except Exception:
            q = None
        return p, q

    if t.endswith("KVA_KVAR"):
        kva = _float_or_none(lv.findtext("KVA"))
        kvar = _float_or_none(lv.findtext("KVAR"))
        if kva is None or kvar is None or kva < abs(kvar):
            return None, kvar
        p = math.sqrt(max(0.0, kva * kva - (kvar * kvar)))
        return p, kvar

    # Fallback
    return _float_or_none(lv.findtext("KW")), _float_or_none(lv.findtext("KVAR"))


def _expand_phases(phase_tag: str) -> List[str]:
    pt = (phase_tag or "").strip().upper()
    if pt in PHASES:
        return [pt]
    if pt in ("AB", "BC", "AC"):
        return list(pt)
    if pt in ("ABC", "", None):
        return ["A", "B", "C"]
    return []


def _parse_spot_and_distributed_loads(txt_path: Path) -> List[Dict[str, Any]]:
    """
    Build observations per declared phase from BOTH SpotLoad and DistributedLoad.
    All bus/ID fields sanitized. We keep zero values; grouping decides visibility.
    """
    root = _read_xml(Path(txt_path))
    out: List[Dict[str, Any]] = []

    for sec in root.findall(".//Sections/Section"):
        section_id = safe_name(sec.findtext("./SectionID"))
        from_bus = safe_name(sec.findtext("./FromNodeID"))
        if not from_bus:
            continue

        # Collect both device types (can be co-located with lines in the same Section)
        devices = list(sec.findall(".//Devices/SpotLoad")) + list(sec.findall(".//Devices/DistributedLoad"))
        if not devices:
            continue

        for dev in devices:
            dev_num = safe_name(dev.findtext("./DeviceNumber"))
            conn_cfg = (dev.findtext("./ConnectionConfiguration") or "").strip()
            # Status & type live under CustomerLoad
            status_txt = (dev.findtext(".//CustomerLoad/ConnectionStatus") or "").strip().lower()
            status = 1 if status_txt == "connected" else 0
            cust_type = (dev.findtext(".//CustomerLoad/CustomerType") or "").strip()
            # Load value type (for ZIP flags)
            lvt = (dev.findtext(".//CustomerLoadModel/LoadValueType") or "").strip().upper()

            load_id = _norm_load_id(dev_num, section_id, from_bus)

            for val in dev.findall(".//CustomerLoadValue"):
                phases = _expand_phases(val.findtext("./Phase"))
                if not phases:
                    continue
                kw, kvar = _kw_kvar_from_value(val)
                share_p = float((kw or 0.0) / len(phases))
                share_q = float((kvar or 0.0) / len(phases))
                for p in phases:
                    out.append({
                        "id": load_id,
                        "bus": from_bus,
                        "phase": p,
                        "conn": conn_cfg,
                        "kw": share_p,
                        "kvar": share_q,
                        "status": status,
                        "cust_type": cust_type,
                        "lvt": lvt,
                        "declared": True,
                    })

    return out


# -----------------------
# Group rows: 1φ / 2φ / 3φ
# -----------------------
def _group_by_device(observations: List[Dict[str, Any]]):
    acc: Dict[str, Dict[str, Dict[str, float]]] = {}
    declared: Dict[str, Set[str]] = {}
    meta: Dict[str, Dict[str, Any]] = {}

    for o in observations:
        lid = o["id"]
        ph = o["phase"]
        acc.setdefault(lid, {})
        acc[lid].setdefault(ph, {"kw": 0.0, "kvar": 0.0})
        acc[lid][ph]["kw"] += o["kw"]
        acc[lid][ph]["kvar"] += o["kvar"]
        declared.setdefault(lid, set()).add(ph)
        # keep the first-seen metadata; if multiple lvt appear, last one wins (rare)
        m = meta.get(lid, {})
        meta[lid] = {
            "bus": o["bus"],
            "conn": o["conn"],
            "status": o["status"],
            "cust_type": o["cust_type"],
            "lvt": o.get("lvt", m.get("lvt")),
        }

    single, two, three = [], [], []
    for lid, phase_map in acc.items():
        nz_phases = [p for p in PHASES
                     if abs(phase_map.get(p, {}).get("kw", 0.0)) > EPS
                     or abs(phase_map.get(p, {}).get("kvar", 0.0)) > EPS]

        phases_used = nz_phases if nz_phases else [p for p in PHASES if p in declared.get(lid, set())]
        if not phases_used:
            continue

        m = meta[lid]
        base = {
            "ID": lid,
            "Status": m["status"],
            "Bus": m["bus"],
            "Conn": m["conn"],
            "CustType": m["cust_type"],
            "LVT": m.get("lvt", ""),
        }

        if len(phases_used) == 1:
            p = phases_used[0]
            row = dict(base)
            row.update({"Phase": p,
                        "P1": phase_map.get(p, {}).get("kw", 0.0),
                        "Q1": phase_map.get(p, {}).get("kvar", 0.0)})
            single.append(row)

        elif len(phases_used) == 2:
            p1, p2 = phases_used[0], phases_used[1]
            row = dict(base)
            row.update({
                "PhasePair": (p1, p2),
                "P1": phase_map.get(p1, {}).get("kw", 0.0),
                "Q1": phase_map.get(p1, {}).get("kvar", 0.0),
                "P2": phase_map.get(p2, {}).get("kw", 0.0),
                "Q2": phase_map.get(p2, {}).get("kvar", 0.0),
            })
            two.append(row)

        else:  # 3 phases
            row = dict(base)
            row.update({
                "P_A": phase_map.get("A", {}).get("kw", 0.0), "Q_A": phase_map.get("A", {}).get("kvar", 0.0),
                "P_B": phase_map.get("B", {}).get("kw", 0.0), "Q_B": phase_map.get("B", {}).get("kvar", 0.0),
                "P_C": phase_map.get("C", {}).get("kw", 0.0), "Q_C": phase_map.get("C", {}).get("kvar", 0.0),
            })
            three.append(row)

    return single, two, three


# -----------------------
# ZIP flags
# -----------------------
def _zip_flags(cust_type: str, load_value_type: str | None = None) -> tuple[int, int, int]:
    """
    Return (Kz, Ki, Kp).
    Prefer explicit LoadValueType when present; otherwise, default to constant power.
    """
    lvt = (load_value_type or "").upper()
    if lvt in ("KW_PF", "KW_KVAR"):
        return 0, 0, 1   # constant power
    if lvt in ("KVA_PF", "KVA_KVAR"):
        return 0, 1, 0   # treat KVA-based as current-type (can be adjusted)
    return 0, 0, 1       # default to constant power


# =======================
# Sheet writer
# =======================
def write_load_sheet(xw, input_path: Path) -> None:
    wb = xw.book
    ws = wb.add_worksheet("Load")
    xw.sheets["Load"] = ws

    # Formats
    bold = wb.add_format({"bold": True})
    link_fmt = wb.add_format({"font_color": "blue", "underline": 1})
    notes_hdr = wb.add_format({"bold": True, "font_color": "yellow", "bg_color": "#595959", "align": "left"})
    notes_txt = wb.add_format({"font_color": "yellow", "bg_color": "#595959"})
    th = wb.add_format({"bold": True, "bottom": 1})
    num2 = wb.add_format({"num_format": "0.00"})
    num0 = wb.add_format({"num_format": "0"})

    # Column widths A..R
    widths = [28, 10, 10, 12, 14, 9, 9, 9, 18, 14, 14, 14, 10, 10, 10, 10, 10, 10]
    for c, w in enumerate(widths):
        ws.set_column(c, c, w)

    # Header
    ws.write(0, 0, "Type", bold)

    # Parse + group (Spot + Distributed)
    obs = _parse_spot_and_distributed_loads(input_path)
    single_rows, two_rows, three_rows = _group_by_device(obs)

    # Voltage map (built on sanitized node names)
    bus_kvll = _build_voltage_map(input_path)

    # ---------- NEW: Determine ACTIVE bus bases from Bus sheet ----------
    bus_rows = extract_bus_data(input_path)
    active_bases: Set[str] = set()
    for row in bus_rows:
        bus_field = str(row.get("Bus", "")).strip()
        if not bus_field or bus_field.startswith("//"):
            continue  # ignore commented Bus rows
        base = _PHASE_SUFFIX_RE.sub("", bus_field)  # strip _a/_b/_c
        if base:
            active_bases.add(base)

    # Anchors
    r = 10
    b1_t, b1_h, b1_e = r, r + 1, r + 2; r = b1_e + 2
    b2_t, b2_h, b2_e = r, r + 1, r + 2; r = b2_e + 2
    b3_t, b3_h, b3_e = r, r + 1, r + 2; r = b3_e + 2
    b4_t, b4_h, b4_first = r, r + 1, r + 2; b4_e = b4_first + len(single_rows); r = b4_e + 2
    b5_t, b5_h, b5_first = r, r + 1, r + 2; b5_e = b5_first + len(two_rows); r = b5_e + 2
    b6_t, b6_h, b6_first = r, r + 1, r + 2; b6_e = b6_first + len(three_rows)

    # Top links
    ws.write_url(1, 0, f"internal:'Load'!A{b1_h+1}:E{b1_e+1}", link_fmt, "PositiveSeqZload")
    ws.write_url(1, 1, f"internal:'Load'!A{b6_h+1}:R{b6_e+1}", link_fmt, "ThreePhaseZIPLoad")
    ws.write_url(2, 0, f"internal:'Load'!A{b2_h+1}:E{b2_e+1}", link_fmt, "PositiveSeqPload")
    ws.write_url(3, 0, f"internal:'Load'!A{b3_h+1}:E{b3_e+1}", link_fmt, "PositiveSeqIload")
    ws.write_url(4, 0, f"internal:'Load'!A{b4_h+1}:L{b4_e+1}", link_fmt, "SinglePhaseZIPLoad")
    ws.write_url(5, 0, f"internal:'Load'!A{b5_h+1}:O{b5_e+1}", link_fmt, "TwoPhaseZIPLoad")

    # Notes
    ws.merge_range(7, 0, 7, 7, "Important notes:", notes_hdr)
    ws.merge_range(8, 0, 8, 7, "Default order of blocks and columns after row 11 must not change", notes_txt)
    ws.merge_range(9, 0, 9, 7, "One empty row between End of each block and the next block is mandatory; otherwise, empty rows are NOT allowed", notes_txt)

    # -------- Block 1: PosSeq Z (empty) --------
    ws.merge_range(b1_t, 0, b1_t, 1, "Positive-Sequence Constant Impedance Load", bold)
    ws.write_url(b1_t, 2, "internal:'Load'!A1", link_fmt, "Go to Type List")
    ws.write_row(b1_h, 0, ["ID", "Status", "Bus", "P (MW)", "Q (MVAr)"], th)
    ws.merge_range(b1_e, 0, b1_e, 2, "End of Positive-Sequence Constant Impedance Load")

    # -------- Block 2: PosSeq P (empty) --------
    ws.merge_range(b2_t, 0, b2_t, 1, "Positive-Sequence Constant Power Load", bold)
    ws.write_url(b2_t, 2, "internal:'Load'!A1", link_fmt, "Go to Type List")
    ws.write_row(b2_h, 0, ["ID", "Status", "Bus", "P (MW)", "Q (MVAr)"], th)
    ws.merge_range(b2_e, 0, b2_e, 2, "End of Positive-Sequence Constant Power Load")

    # -------- Block 3: PosSeq I (empty) --------
    ws.merge_range(b3_t, 0, b3_t, 1, "Positive-Sequence Constant Current Load", bold)
    ws.write_url(b3_t, 2, "internal:'Load'!A1", link_fmt, "Go to Type List")
    ws.write_row(b3_h, 0, ["ID", "Status", "Bus", "P (MW)", "Q (MVAr)"], th)
    ws.merge_range(b3_e, 0, b3_e, 2, "End of Positive-Sequence Constant Current Load")

    # -------- Block 4: Single-Phase ZIP --------
    ws.merge_range(b4_t, 0, b4_t, 1, "Single-Phase ZIP Load", bold)
    ws.write_url(b4_t, 2, "internal:'Load'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b4_h, 0,
        ["ID", "Status", "V (kV)", "Bandwidth (pu)", "Conn. type", "K_z", "K_i", "K_p",
         "Use initial voltage?", "Bus1", "P1 (kW)", "Q1 (kVAr)"],
        th,
    )
    rcur = b4_first
    for row in single_rows:
        kz, ki, kp = _zip_flags(row["CustType"], row.get("LVT"))
        bus = row["Bus"]
        vkv = bus_kvll.get(bus, bus_kvll["_default_"])
        conn = (row["Conn"] or "").lower()

        # NEW: comment out this row if its bus base is NOT active on Bus sheet
        is_active_bus = (bus in active_bases)
        id_out = row["ID"] if is_active_bus else f"//{row['ID']}"

        ws.write(rcur, 0, id_out)
        ws.write_number(rcur, 1, row["Status"], num0)
        ws.write_number(rcur, 2, vkv, num2)
        ws.write_number(rcur, 3, 0.2, num2)
        ws.write(rcur, 4, "wye" if conn.startswith("y") else "delta" if conn.startswith("d") else "")
        ws.write_number(rcur, 5, kz, num0)
        ws.write_number(rcur, 6, ki, num0)
        ws.write_number(rcur, 7, kp, num0)
        ws.write_number(rcur, 8, 0, num0)
        ws.write(rcur, 9, f"{bus}{PHASE_SUFFIX.get(row['Phase'], '')}")
        ws.write_number(rcur, 10, row["P1"], num0)
        ws.write_number(rcur, 11, row["Q1"], num0)
        rcur += 1
    ws.merge_range(b4_e, 0, b4_e, 1, "End of Single-Phase ZIP Load")

    # -------- Block 5: Two-Phase ZIP --------
    ws.merge_range(b5_t, 0, b5_t, 1, "Two-Phase ZIP Load", bold)
    ws.write_url(b5_t, 2, "internal:'Load'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b5_h, 0,
        ["ID", "Status", "V (kV)", "Bandwidth (pu)", "Conn. type", "K_z", "K_i", "K_p",
         "Use initial voltage?", "Bus1", "Bus2", "P1(kW)", "Q1(kVAr)", "P2 (kW)", "Q2 (kVAr)"],
        th,
    )
    rcur = b5_first
    for row in two_rows:
        kz, ki, kp = _zip_flags(row["CustType"], row.get("LVT"))
        bus = row["Bus"]
        vkv = bus_kvll.get(bus, bus_kvll["_default_"])
        p1, p2 = row["PhasePair"]
        conn = (row["Conn"] or "").lower()

        is_active_bus = (bus in active_bases)
        id_out = row["ID"] if is_active_bus else f"//{row['ID']}"

        ws.write(rcur, 0, id_out); ws.write_number(rcur, 1, row["Status"], num0)
        ws.write_number(rcur, 2, vkv, num2)
        ws.write_number(rcur, 3, 0.2, num2)
        ws.write(rcur, 4, "wye" if conn.startswith("y") else "delta" if conn.startswith("d") else "")
        ws.write_number(rcur, 5, kz, num0); ws.write_number(rcur, 6, ki, num0); ws.write_number(rcur, 7, kp, num0)
        ws.write_number(rcur, 8, 0, num0)
        ws.write(rcur, 9,  f"{bus}{PHASE_SUFFIX[p1]}"); ws.write(rcur,10, f"{bus}{PHASE_SUFFIX[p2]}")
        ws.write_number(rcur,11, row["P1"], num0); ws.write_number(rcur,12, row["Q1"], num0)
        ws.write_number(rcur,13, row["P2"], num0); ws.write_number(rcur,14, row["Q2"], num0)
        rcur += 1
    ws.merge_range(b5_e, 0, b5_e, 1, "End of Two-Phase ZIP Load")

    # -------- Block 6: Three-Phase ZIP --------
    ws.merge_range(b6_t, 0, b6_t, 1, "Three-Phase ZIP Load", bold)
    ws.write_url(b6_t, 2, "internal:'Load'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b6_h, 0,
        ["ID", "Status", "V (kV)", "Bandwidth (pu)", "Conn. type", "K_z", "K_i", "K_p",
         "Use initial voltage?", "Bus1", "Bus2", "Bus3",
         "P1(kW)", "Q1(kVAr)", "P2 (kW)", "Q2 (kVAr)", "P3 (kW)", "Q3 (kVAr)"],
        th,
    )
    rcur = b6_first
    for row in three_rows:
        kz, ki, kp = _zip_flags(row["CustType"], row.get("LVT"))
        bus = row["Bus"]
        vkv = bus_kvll.get(bus, bus_kvll["_default_"])
        conn = (row["Conn"] or "").lower()

        is_active_bus = (bus in active_bases)
        id_out = row["ID"] if is_active_bus else f"//{row['ID']}"

        ws.write(rcur, 0, id_out); ws.write_number(rcur, 1, row["Status"], num0)
        ws.write_number(rcur, 2, vkv, num2)
        ws.write_number(rcur, 3, 0.2, num2)
        ws.write(rcur, 4, "wye" if conn.startswith("y") else "delta" if conn.startswith("d") else "")
        ws.write_number(rcur, 5, kz, num0); ws.write_number(rcur, 6, ki, num0); ws.write_number(rcur, 7, kp, num0)
        ws.write_number(rcur, 8, 0, num0)
        ws.write(rcur, 9,  f"{bus}_a"); ws.write(rcur,10, f"{bus}_b"); ws.write(rcur,11, f"{bus}_c")
        ws.write_number(rcur,12, row["P_A"], num0); ws.write_number(rcur,13, row["Q_A"], num0)
        ws.write_number(rcur,14, row["P_B"], num0); ws.write_number(rcur,15, row["Q_B"], num0)
        ws.write_number(rcur,16, row["P_C"], num0); ws.write_number(rcur,17, row["Q_C"], num0)
        rcur += 1
    ws.merge_range(b6_e, 0, b6_e, 1, "End of Three-Phase ZIP Load")
