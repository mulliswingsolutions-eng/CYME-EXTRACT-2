# Modules/Load.py
from __future__ import annotations
from pathlib import Path
import re
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple, Any

# -----------------------
# Constants
# -----------------------
PHASES = ("A", "B", "C")
PHASE_SUFFIX = {"A": "_a", "B": "_b", "C": "_c"}


# =======================
# XML helpers
# =======================
def _digits(s: str) -> str:
    return "".join(ch for ch in (s or "") if ch.isdigit())


def _get_default_kvll(root: ET.Element) -> float:
    """
    Read feeder line-to-line base (kV) from the Equivalent Source KVLL.
    """
    eq = root.find(".//Sources/Source/EquivalentSourceModels/EquivalentSourceModel/EquivalentSource")
    if eq is not None:
        try:
            return float(eq.findtext("KVLL", "0") or "0")
        except Exception:
            pass
    # Fallback if needed (shouldn't happen on 13-bus)
    return 0


def _transformer_db_kvll(root: ET.Element) -> Dict[str, float]:
    """
    Map TransformerDB DeviceID -> SecondaryVoltage (kV LL).
    """
    out: Dict[str, float] = {}
    for tdb in root.findall(".//TransformerDB"):
        dev = (tdb.findtext("EquipmentID") or "").strip()
        if not dev:
            continue
        try:
            sec = float(tdb.findtext("SecondaryVoltage", "") or "nan")
        except Exception:
            sec = float("nan")
        if sec == sec:  # not NaN
            out[dev] = sec
    return out


def _secondary_bus_by_transformer(root: ET.Element) -> Dict[str, str]:
    """
    Map Transformer 'DeviceID' used in sections -> secondary bus ID.

    We consider the node opposite to NormalFeedingNodeID as the secondary.
    (If missing, we assume ToNodeID is secondary.)
    """
    result: Dict[str, str] = {}
    for sec in root.findall(".//Section"):
        xf = sec.find(".//Transformer")
        if xf is None:
            continue
        dev = (xf.findtext("DeviceID") or "").strip()
        if not dev:
            continue

        from_node = (sec.findtext("FromNodeID") or "").strip()
        to_node = (sec.findtext("ToNodeID") or "").strip()
        normal_feed = (xf.findtext("NormalFeedingNodeID") or "").strip()

        if normal_feed and normal_feed == from_node and to_node:
            secondary = to_node
        elif normal_feed and normal_feed == to_node and from_node:
            secondary = from_node
        else:
            # best-effort fallback
            secondary = to_node or from_node

        if dev and secondary:
            result[dev] = secondary
    return result


def _kvll_map_for_buses(txt_path: Path) -> Dict[str, float]:
    """
    Build bus -> kVLL map:
      - default from equivalent source KVLL (4.16 kV here)
      - overrides for transformer secondary buses from TransformerDB.SecondaryVoltage
        using the transformer sections to locate the secondary bus.
    """
    root = ET.fromstring(txt_path.read_text(encoding="utf-8", errors="ignore"))
    default_kvll = _get_default_kvll(root)
    dev_to_sec_kv = _transformer_db_kvll(root)
    dev_to_sec_bus = _secondary_bus_by_transformer(root)

    bus_kv: Dict[str, float] = {}
    for dev, bus in dev_to_sec_bus.items():
        sec_kv = dev_to_sec_kv.get(dev)
        if sec_kv:
            bus_kv[bus] = sec_kv

    # default applies implicitly for any bus not in overrides
    bus_kv["_default_"] = default_kvll
    return bus_kv


# -----------------------
# Helpers: ID normalize
# -----------------------
def _norm_load_id(device_number: str, section_id: str, from_node_id: str) -> str:
    for src in (device_number, section_id, from_node_id):
        if not src:
            continue
        m = re.search(r"(\d+)", src)
        if m:
            return f"LD_{m.group(1)}"
    return (device_number or section_id or from_node_id or "LD_?").replace(" ", "_")


# -----------------------
# Parse every SpotLoad
# -----------------------
def _parse_spot_loads_all(txt_path: Path) -> List[Dict[str, Any]]:
    """
    Flat list of observations (one per phase value found).
    Keys: id, bus, phase, conn, kw, kvar, status, cust_type
    """
    xml_text = txt_path.read_text(encoding="utf-8", errors="ignore")
    root = ET.fromstring(xml_text)

    out: List[Dict[str, Any]] = []
    for sec in root.findall(".//Section"):
        section_id = (sec.findtext("./SectionID") or "").strip()
        from_bus = (sec.findtext("./FromNodeID") or "").strip()

        spot = sec.find(".//SpotLoad")
        if spot is None:
            continue

        dev_num = (spot.findtext("./DeviceNumber") or "").strip()
        conn_cfg = (spot.findtext("./ConnectionConfiguration") or "").strip()
        cust_type = (spot.findtext(".//CustomerLoad/CustomerType") or "").strip().upper()
        status_txt = (spot.findtext(".//CustomerLoad/ConnectionStatus") or "").strip().lower()
        status = 1 if status_txt == "connected" else 0

        load_id = _norm_load_id(dev_num, section_id, from_bus)

        for val in spot.findall(".//CustomerLoadValue"):
            ph = (val.findtext("./Phase") or "").strip().upper()
            if ph not in PHASES:
                continue
            kw_txt = val.findtext(".//KW")
            kvar_txt = val.findtext(".//KVAR")
            try:
                kw = float(kw_txt) if kw_txt is not None else None
                kvar = float(kvar_txt) if kvar_txt is not None else None
            except Exception:
                kw, kvar = None, None
            if kw is None or kvar is None:
                continue

            out.append(
                {
                    "id": load_id,
                    "bus": from_bus,
                    "phase": ph,
                    "conn": conn_cfg,
                    "kw": kw,
                    "kvar": kvar,
                    "status": status,
                    "cust_type": cust_type,
                }
            )

    return out


# -----------------------
# Group: 1φ / 2φ / 3φ
# -----------------------
def _group_by_device(observations: List[Dict[str, Any]]):
    acc: Dict[str, Dict[str, Dict[str, float]]] = {}
    meta: Dict[str, Dict[str, Any]] = {}

    for o in observations:
        lid = o["id"]
        ph = o["phase"]
        acc.setdefault(lid, {})
        acc[lid].setdefault(ph, {"kw": 0.0, "kvar": 0.0})
        acc[lid][ph]["kw"] += o["kw"]
        acc[lid][ph]["kvar"] += o["kvar"]

        meta[lid] = {
            "bus": o["bus"],
            "conn": o["conn"],
            "status": o["status"],
            "cust_type": o["cust_type"],
        }

    single, two, three = [], [], []

    for lid, phase_map in acc.items():
        phases = sorted(phase_map.keys(), key=lambda p: PHASES.index(p))
        m = meta[lid]
        entry = {
            "ID": lid,
            "Status": m["status"],
            "Bus": m["bus"],
            "Conn": m["conn"],
            "CustType": m["cust_type"],
        }

        if len(phases) == 1:
            ph = phases[0]
            row = dict(entry)
            row.update({"Phase": ph, "P1": phase_map[ph]["kw"], "Q1": phase_map[ph]["kvar"]})
            single.append(row)
        elif len(phases) == 2:
            p1, p2 = phases[0], phases[1]
            row = dict(entry)
            row.update(
                {
                    "PhasePair": (p1, p2),
                    "P1": phase_map[p1]["kw"],
                    "Q1": phase_map[p1]["kvar"],
                    "P2": phase_map[p2]["kw"],
                    "Q2": phase_map[p2]["kvar"],
                }
            )
            two.append(row)
        elif len(phases) == 3:
            row = dict(entry)
            row.update(
                {
                    "P_A": phase_map["A"]["kw"],
                    "Q_A": phase_map["A"]["kvar"],
                    "P_B": phase_map["B"]["kw"],
                    "Q_B": phase_map["B"]["kvar"],
                    "P_C": phase_map["C"]["kw"],
                    "Q_C": phase_map["C"]["kvar"],
                }
            )
            three.append(row)

    return single, two, three


# -----------------------
# Kz / Ki / Kp flags
# -----------------------
def _zip_flags(cust_type: str) -> Tuple[int, int, int]:
    t = (cust_type or "").upper()
    if t.startswith("Z"):
        return 1, 0, 0
    if t.startswith("I"):
        return 0, 1, 0
    return 0, 0, 1  # PQ or anything else -> constant power


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

    # Type header + top links
    ws.write(0, 0, "Type", bold)

    # Parse + group
    xml_text = input_path.read_text(encoding="utf-8", errors="ignore")
    root = ET.fromstring(xml_text)
    obs = _parse_spot_loads_all(input_path)
    single_rows, two_rows, three_rows = _group_by_device(obs)

    # Build bus -> V(kV, LL) map (default 4.16, with 634 -> 0.48 via transformer DB)
    bus_kvll = _kvll_map_for_buses(input_path)

    # Anchor rows
    r = 10
    b1_t, b1_h, b1_e = r, r + 1, r + 2; r = b1_e + 2
    b2_t, b2_h, b2_e = r, r + 1, r + 2; r = b2_e + 2
    b3_t, b3_h, b3_e = r, r + 1, r + 2; r = b3_e + 2
    b4_t, b4_h, b4_first = r, r + 1, r + 2; b4_e = b4_first + len(single_rows); r = b4_e + 2
    b5_t, b5_h, b5_first = r, r + 1, r + 2; b5_e = b5_first + len(two_rows); r = b5_e + 2
    b6_t, b6_h, b6_first = r, r + 1, r + 2; b6_e = b6_first + len(three_rows)

    # Top link ranges (with ThreePhaseZIPLoad in B2)
    ws.write_url(1, 0, f"internal:'Load'!A{b1_h+1}:E{b1_e+1}", link_fmt, "PositiveSeqZload")        # A2
    ws.write_url(1, 1, f"internal:'Load'!A{b6_h+1}:R{b6_e+1}", link_fmt, "ThreePhaseZIPLoad")       # B2
    ws.write_url(2, 0, f"internal:'Load'!A{b2_h+1}:E{b2_e+1}", link_fmt, "PositiveSeqPload")        # A3
    ws.write_url(3, 0, f"internal:'Load'!A{b3_h+1}:E{b3_e+1}", link_fmt, "PositiveSeqIload")        # A4
    ws.write_url(4, 0, f"internal:'Load'!A{b4_h+1}:L{b4_e+1}", link_fmt, "SinglePhaseZIPLoad")      # A5
    ws.write_url(5, 0, f"internal:'Load'!A{b5_h+1}:O{b5_e+1}", link_fmt, "TwoPhaseZIPLoad")         # A6

    # Notes (rows 8–10), merged A:H
    ws.merge_range(7, 0, 7, 7, "Important notes:", notes_hdr)
    ws.merge_range(8, 0, 8, 7, "Default order of blocks and columns after row 11 must not change", notes_txt)
    ws.merge_range(9, 0, 9, 7, "One empty row between End of each block and the next block is mandatory; otherwise, empty rows are NOT allowed", notes_txt)

    # -------- Block 1: PosSeq Z (empty) --------
    ws.merge_range(b1_t, 0, b1_t, 3, "Positive-Sequence Constant Impedance Load", bold)
    ws.write_url(b1_t, 4, "internal:'Load'!A1", link_fmt, "Go to Type List")
    ws.write_row(b1_h, 0, ["ID", "Status", "Bus", "P (MW)", "Q (MVAr)"], th)
    ws.merge_range(b1_e, 0, b1_e, 3, "End of Positive-Sequence Constant Impedance Load")

    # -------- Block 2: PosSeq P (empty) --------
    ws.merge_range(b2_t, 0, b2_t, 3, "Positive-Sequence Constant Power Load", bold)
    ws.write_url(b2_t, 4, "internal:'Load'!A1", link_fmt, "Go to Type List")
    ws.write_row(b2_h, 0, ["ID", "Status", "Bus", "P (MW)", "Q (MVAr)"], th)
    ws.merge_range(b2_e, 0, b2_e, 3, "End of Positive-Sequence Constant Power Load")

    # -------- Block 3: PosSeq I (empty) --------
    ws.merge_range(b3_t, 0, b3_t, 3, "Positive-Sequence Constant Current Load", bold)
    ws.write_url(b3_t, 4, "internal:'Load'!A1", link_fmt, "Go to Type List")
    ws.write_row(b3_h, 0, ["ID", "Status", "Bus", "P (MW)", "Q (MVAr)"], th)
    ws.merge_range(b3_e, 0, b3_e, 3, "End of Positive-Sequence Constant Current Load")

    # -------- Block 4: Single-Phase ZIP (parsed) --------
    ws.merge_range(b4_t, 0, b4_t, 3, "Single-Phase ZIP Load", bold)
    ws.write_url(b4_t, 4, "internal:'Load'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b4_h, 0,
        ["ID", "Status", "V (kV)", "Bandwidth (pu)", "Conn. type", "K_z", "K_i", "K_p",
         "Use initial voltage?", "Bus1", "P1 (kW)", "Q1 (kVAr)"],
        th,
    )
    rcur = b4_first
    for row in single_rows:
        kz, ki, kp = _zip_flags(row["CustType"])
        bus = row["Bus"]
        vkv = bus_kvll.get(bus, bus_kvll["_default_"])

        conn = (row["Conn"] or "").lower()
        ws.write(rcur, 0, row["ID"])
        ws.write_number(rcur, 1, row["Status"], num0)
        ws.write_number(rcur, 2, vkv, num2)
        ws.write_number(rcur, 3, 0.2, num2)
        ws.write(rcur, 4, "wye" if conn.startswith("y") else "delta" if conn.startswith("d") else "")
        ws.write_number(rcur, 5, kz, num0)
        ws.write_number(rcur, 6, ki, num0)
        ws.write_number(rcur, 7, kp, num0)
        ws.write_number(rcur, 8, 0, num0)  # Use initial voltage?
        ws.write(rcur, 9, f"{bus}{PHASE_SUFFIX.get(row['Phase'], '')}")
        ws.write_number(rcur, 10, row["P1"], num0)
        ws.write_number(rcur, 11, row["Q1"], num0)
        rcur += 1
    ws.merge_range(b4_e, 0, b4_e, 3, "End of SinglePhase ZIP Load")

    # -------- Block 5: Two-Phase ZIP (parsed) --------
    ws.merge_range(b5_t, 0, b5_t, 3, "Two-Phase ZIP Load", bold)
    ws.write_url(b5_t, 4, "internal:'Load'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b5_h, 0,
        ["ID", "Status", "V (kV)", "Bandwidth (pu)", "Conn. type", "K_z", "K_i", "K_p",
         "Use initial voltage?", "Bus1", "Bus2", "P1(kW)", "Q1(kVAr)", "P2 (kW)", "Q2 (kVAr)"],
        th,
    )
    rcur = b5_first
    for row in two_rows:
        kz, ki, kp = _zip_flags(row["CustType"])
        bus = row["Bus"]
        vkv = bus_kvll.get(bus, bus_kvll["_default_"])
        p1, p2 = row["PhasePair"]
        conn = (row["Conn"] or "").lower()

        ws.write(rcur, 0, row["ID"]); ws.write_number(rcur, 1, row["Status"], num0)
        ws.write_number(rcur, 2, vkv, num2)
        ws.write_number(rcur, 3, 0.2, num2)
        ws.write(rcur, 4, "wye" if conn.startswith("y") else "delta" if conn.startswith("d") else "")
        ws.write_number(rcur, 5, kz, num0); ws.write_number(rcur, 6, ki, num0); ws.write_number(rcur, 7, kp, num0)
        ws.write_number(rcur, 8, 0, num0)
        ws.write(rcur, 9,  f"{bus}{PHASE_SUFFIX[p1]}"); ws.write(rcur,10, f"{bus}{PHASE_SUFFIX[p2]}")
        ws.write_number(rcur,11, row["P1"], num0); ws.write_number(rcur,12, row["Q1"], num0)
        ws.write_number(rcur,13, row["P2"], num0); ws.write_number(rcur,14, row["Q2"], num0)
        rcur += 1
    ws.merge_range(b5_e, 0, b5_e, 3, "End of TwoPhase ZIP Load")

    # -------- Block 6: Three-Phase ZIP (parsed) --------
    ws.merge_range(b6_t, 0, b6_t, 3, "Three-Phase ZIP Load", bold)
    ws.write_url(b6_t, 4, "internal:'Load'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b6_h, 0,
        ["ID", "Status", "V (kV)", "Bandwidth (pu)", "Conn. type", "K_z", "K_i", "K_p",
         "Use initial voltage?", "Bus1", "Bus2", "Bus3",
         "P1(kW)", "Q1(kVAr)", "P2 (kW)", "Q2 (kVAr)", "P3 (kW)", "Q3 (kVAr)"],
        th,
    )
    rcur = b6_first
    for row in three_rows:
        kz, ki, kp = _zip_flags(row["CustType"])
        bus = row["Bus"]
        vkv = bus_kvll.get(bus, bus_kvll["_default_"])
        conn = (row["Conn"] or "").lower()

        ws.write(rcur, 0, row["ID"]); ws.write_number(rcur, 1, row["Status"], num0)
        ws.write_number(rcur, 2, vkv, num2)
        ws.write_number(rcur, 3, 0.2, num2)
        ws.write(rcur, 4, "wye" if conn.startswith("y") else "delta" if conn.startswith("d") else "")
        ws.write_number(rcur, 5, kz, num0); ws.write_number(rcur, 6, ki, num0); ws.write_number(rcur, 7, kp, num0)
        ws.write_number(rcur, 8, 0, num0)
        ws.write(rcur, 9,  f"{bus}{PHASE_SUFFIX['A']}"); ws.write(rcur,10, f"{bus}{PHASE_SUFFIX['B']}"); ws.write(rcur,11, f"{bus}{PHASE_SUFFIX['C']}")
        ws.write_number(rcur,12, row["P_A"], num0); ws.write_number(rcur,13, row["Q_A"], num0)
        ws.write_number(rcur,14, row["P_B"], num0); ws.write_number(rcur,15, row["Q_B"], num0)
        ws.write_number(rcur,16, row["P_C"], num0); ws.write_number(rcur,17, row["Q_C"], num0)
        rcur += 1
    ws.merge_range(b6_e, 0, b6_e, 3, "End of Three-Phase ZIP Load")
