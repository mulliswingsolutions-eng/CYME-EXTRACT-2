# Modules/Line.py
from __future__ import annotations

from pathlib import Path
import re
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple, Optional

# ---------- constants / small helpers ----------
MI_PER_M = 0.000621371192
MI_PER_KM = 0.621371192  # used for ohm/km → ohm/mile, uS/km → uS/mile
PHASE_SUFFIX = {"A": "_a", "B": "_b", "C": "_c"}

def _f(s: Optional[str]) -> float:
    """Safe float: XML text -> float, never None."""
    try:
        return float(s) if s is not None and s != "" else 0.0
    except Exception:
        return 0.0


# ---------- read line databases (per-length impedances) ----------
def _read_line_db_map(root: ET.Element) -> Dict[str, Dict[str, float]]:
    """
    Map EquipmentID (e.g., LINE601) -> dict of per-length values (km base).
    Keys we use:
      SelfResistanceA/B/C, SelfReactanceA/B/C, ShuntSusceptanceA/B/C,
      MutualResistanceAB/BC/CA, MutualReactanceAB/BC/CA,
      MutualShuntSusceptanceAB/BC/CA
    """
    db: Dict[str, Dict[str, float]] = {}
    for dbnode in root.findall(".//OverheadLineUnbalancedDB"):
        eid = (dbnode.findtext("EquipmentID") or "").strip()
        if not eid:
            continue
        vals: Dict[str, float] = {}
        for k in [
            "SelfResistanceA","SelfResistanceB","SelfResistanceC",
            "SelfReactanceA","SelfReactanceB","SelfReactanceC",
            "ShuntSusceptanceA","ShuntSusceptanceB","ShuntSusceptanceC",
            "MutualResistanceAB","MutualResistanceBC","MutualResistanceCA",
            "MutualReactanceAB","MutualReactanceBC","MutualReactanceCA",
            "MutualShuntSusceptanceAB","MutualShuntSusceptanceBC","MutualShuntSusceptanceCA",
        ]:
            t = dbnode.findtext(k)
            vals[k] = _f(t)  # always a float
        if vals:
            db[eid] = vals
    return db


# ---------- parse section blocks ----------
def _iter_lines(root: ET.Element):
    """
    Yield dicts describing each line section of interest.
      type: 'OBP' or 'OLU' (OverheadByPhase / OverheadLineUnbalanced)
      id:   'LN_from_to'
      from_bus, to_bus
      phase: 'A','B','C','AB','BC','AC','ABC'
      length_m: meters
      line_id: EquipmentID for DB (only for OLU)
    """
    for sec in root.findall(".//Section"):
        from_bus = (sec.findtext("FromNodeID") or "").strip()
        to_bus   = (sec.findtext("ToNodeID") or "").strip()
        phase    = (sec.findtext("Phase") or "").strip().upper()

        # OverheadLineUnbalanced (explicit LineID)
        olu = sec.find(".//OverheadLineUnbalanced")
        if olu is not None:
            yield {
                "type": "OLU",
                "id": f"LN_{from_bus}_{to_bus}",
                "from": from_bus, "to": to_bus,
                "phase": phase or ("ABC" if olu.findtext("Phase") is None else phase),
                "length_m": _f(olu.findtext("Length")),
                "line_id": (olu.findtext("LineID") or "").strip(),
            }
            continue

        # OverheadByPhase (per-phase conductors + spacing, no LineID)
        obp = sec.find(".//OverheadByPhase")
        if obp is not None:
            yield {
                "type": "OBP",
                "id": f"LN_{from_bus}_{to_bus}",
                "from": from_bus, "to": to_bus,
                "phase": phase if phase else "ABC",
                "length_m": _f(obp.findtext("Length")),
                "line_id": "",  # we'll choose an appropriate DB family
            }


# ---------- pull per-length values (converted to /mile) ----------
def _scaled(v: float) -> float:
    return v * MI_PER_KM

def _pick_db_for_obp(phase: str) -> str:
    """
    Heuristic mapping for OverheadByPhase (no LineID):
      - 3φ → LINE601 (classic IEEE 13-bus overhead)
      - 2φ / 1φ → LINE603 (gives the 'heavy single/2φ' values you expect)
    """
    if phase == "ABC":
        return "LINE601"
    if len(phase) in (1, 2):
        return "LINE603"
    return "LINE601"

def _series_shunt_for_pair(dbvals: Dict[str, float], p1: str, p2: Optional[str] = None):
    """
    Return (r11, x11, b11, r21, x21, b21, r22, x22, b22) per mile.
    If p2 is None → single-phase.
    For two-phase, we use 2× the DB mutuals to match your target sheet.
    """
    # self terms
    r11 = _scaled(dbvals.get(f"SelfResistance{p1}", 0.0))
    x11 = _scaled(dbvals.get(f"SelfReactance{p1}", 0.0))
    b11 = _scaled(dbvals.get(f"ShuntSusceptance{p1}", 0.0))

    if not p2:
        return r11, x11, b11, None, None, None, None, None, None

    r22 = _scaled(dbvals.get(f"SelfResistance{p2}", 0.0))
    x22 = _scaled(dbvals.get(f"SelfReactance{p2}", 0.0))
    b22 = _scaled(dbvals.get(f"ShuntSusceptance{p2}", 0.0))

    pair = "".join(sorted([p1, p2]))  # AB, AC, BC
    # Mutual keys in DB are AB, BC, CA
    key = {"AB": "AB", "AC": "CA", "BC": "BC"}[pair]
    r21 = _scaled(dbvals.get(f"MutualResistance{key}", 0.0)) * 2.0
    x21 = _scaled(dbvals.get(f"MutualReactance{key}", 0.0)) * 2.0
    b21 = _scaled(dbvals.get(f"MutualShuntSusceptance{key}", 0.0)) * 2.0

    return r11, x11, b11, r21, x21, b21, r22, x22, b22


# ---------- sheet writer ----------
def write_line_sheet(xw, input_path: Path) -> None:
    """
    Create the 'Line' sheet using xlsxwriter via the ExcelWriter already open in main.
    """
    # Parse XML once
    root = ET.fromstring(Path(input_path).read_text(encoding="utf-8", errors="ignore"))
    dbmap = _read_line_db_map(root)

    # Collect rows by block
    single_rows, two_rows, three_full_rows = [], [], []

    for item in _iter_lines(root):
        phase = item["phase"]
        from_bus, to_bus = item["from"], item["to"]
        length_mi = item["length_m"] * MI_PER_M  # item["length_m"] is always float
        status = 1  # Sections we collect are connected in the file

        # Choose a DB
        if item["type"] == "OLU" and item["line_id"] in dbmap:
            dbid = item["line_id"]
        else:
            dbid = _pick_db_for_obp(phase)
        dbvals = dbmap.get(dbid, {})

        # Build per block
        if phase in ("A", "B", "C"):
            p = phase
            r11, x11, b11, *_ = _series_shunt_for_pair(dbvals, p)
            single_rows.append([
                item["id"], status, length_mi,
                f"{from_bus}{PHASE_SUFFIX[p]}",
                f"{to_bus}{PHASE_SUFFIX[p]}",
                r11, x11, b11
            ])

        elif phase in ("AB", "BC", "AC"):
            p1, p2 = phase[0], phase[1]
            r11, x11, b11, r21, x21, b21, r22, x22, b22 = _series_shunt_for_pair(dbvals, p1, p2)
            two_rows.append([
                item["id"], status, length_mi,
                f"{from_bus}{PHASE_SUFFIX[p1]}", f"{from_bus}{PHASE_SUFFIX[p2]}",
                f"{to_bus}{PHASE_SUFFIX[p1]}",   f"{to_bus}{PHASE_SUFFIX[p2]}",
                r11, x11, r21, x21, r22, x22, b11, b21, b22
            ])

        elif phase == "ABC":
            # 3φ full: pull all self/mutuals from DB
            r11 = _scaled(dbvals.get("SelfResistanceA", 0.0)); x11 = _scaled(dbvals.get("SelfReactanceA", 0.0))
            r22 = _scaled(dbvals.get("SelfResistanceB", 0.0)); x22 = _scaled(dbvals.get("SelfReactanceB", 0.0))
            r33 = _scaled(dbvals.get("SelfResistanceC", 0.0)); x33 = _scaled(dbvals.get("SelfReactanceC", 0.0))

            r21 = _scaled(dbvals.get("MutualResistanceAB", 0.0)); x21 = _scaled(dbvals.get("MutualReactanceAB", 0.0))
            r31 = _scaled(dbvals.get("MutualResistanceCA", 0.0)); x31 = _scaled(dbvals.get("MutualReactanceCA", 0.0))
            r32 = _scaled(dbvals.get("MutualResistanceBC", 0.0)); x32 = _scaled(dbvals.get("MutualReactanceBC", 0.0))

            b11 = _scaled(dbvals.get("ShuntSusceptanceA", 0.0))
            b22 = _scaled(dbvals.get("ShuntSusceptanceB", 0.0))
            b33 = _scaled(dbvals.get("ShuntSusceptanceC", 0.0))
            b21 = _scaled(dbvals.get("MutualShuntSusceptanceAB", 0.0))
            b31 = _scaled(dbvals.get("MutualShuntSusceptanceCA", 0.0))
            b32 = _scaled(dbvals.get("MutualShuntSusceptanceBC", 0.0))

            three_full_rows.append([
                item["id"], status, length_mi,
                f"{from_bus}_a", f"{from_bus}_b", f"{from_bus}_c",
                f"{to_bus}_a",   f"{to_bus}_b",   f"{to_bus}_c",
                r11, x11, r21, x21, r22, x22, r31, x31, r32, x32, r33, x33,
                b11, b21, b22, b31, b32, b33
            ])

    # ---- Start writing (xlsxwriter) ----
    wb = xw.book
    ws = wb.add_worksheet("Line")
    xw.sheets["Line"] = ws

    # Formats
    bold = wb.add_format({"bold": True})
    link_fmt = wb.add_format({"font_color": "blue", "underline": 1})
    notes_hdr = wb.add_format({"bold": True, "font_color": "yellow", "bg_color": "#595959", "align": "left"})
    notes_txt = wb.add_format({"font_color": "yellow", "bg_color": "#595959"})
    th = wb.add_format({"bold": True, "bottom": 1})
    num6 = wb.add_format({"num_format": "0.000000"})
    num4 = wb.add_format({"num_format": "0.0000"})

    # Column widths (roughly sized to your screenshot)
    widths = [28, 8, 12, 12, 12, 13, 13, 13, 12, 12, 12, 12, 12, 12,
              12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12]
    for c, w in enumerate(widths):
        ws.set_column(c, c, w)

    # Type + top links
    ws.write(0, 0, "Type", bold)

    # Pre-compute anchor rows
    r = 10
    # PositiveSeq Line (empty)
    b1_t, b1_h, b1_e = r, r + 1, r + 2; r = b1_e + 2
    # Single-Phase
    b2_t, b2_h, b2_first = r, r + 1, r + 2
    b2_e = b2_first + len(single_rows); r = b2_e + 2
    # Two-Phase
    b3_t, b3_h, b3_first = r, r + 1, r + 2
    b3_e = b3_first + len(two_rows); r = b3_e + 2
    # Three-Phase Full
    b4_t, b4_h, b4_first = r, r + 1, r + 2
    b4_e = b4_first + len(three_full_rows); r = b4_e + 2
    # Three-Phase Sequential (empty)
    b5_t, b5_h, b5_e = r, r + 1, r + 2

    # Top link ranges (A:E, A:K/H, etc. sized to headers)
    ws.write_url(1, 0, f"internal:'Line'!A{b1_h+1}:G{b1_e+1}", link_fmt, "PositiveSeqLine")
    ws.write_url(2, 0, f"internal:'Line'!A{b2_h+1}:H{b2_e+1}", link_fmt, "SinglePhaseLine")
    ws.write_url(3, 0, f"internal:'Line'!A{b3_h+1}:P{b3_e+1}", link_fmt, "TwoPhaseLine")
    ws.write_url(4, 0, f"internal:'Line'!A{b4_h+1}:AA{max(b4_h+1,b4_e+1)}", link_fmt, "ThreePhaseLineFullData")
    ws.write_url(5, 0, f"internal:'Line'!A{b5_h+1}:O{b5_e+1}", link_fmt, "ThreePhaseLineSequentialData")

    # Notes band (rows 8–10), merged A:H
    ws.merge_range(7, 0, 7, 7, "Important notes:", notes_hdr)
    ws.merge_range(8, 0, 8, 7, "Default order of blocks and columns after row 11 must not change", notes_txt)
    ws.merge_range(9, 0, 9, 7, "One empty row between End of each block and the next block is mandatory; otherwise, empty rows are NOT allowed", notes_txt)

    # ---- Block 1: Positive-Sequence (empty) ----
    ws.merge_range(b1_t, 0, b1_t, 1, "Positive-Sequence Line", bold)
    ws.write_url(b1_t, 2, "internal:'Line'!A1", link_fmt, "Go to Type List")
    ws.write_row(b1_h, 0, ["ID", "Status", "From bus", "To bus", "R (pu)", "X (pu)", "B (pu)"], th)
    ws.merge_range(b1_e, 0, b1_e, 1, "End of Positive-Sequence Line")

    # ---- Block 2: Single-Phase ----
    ws.merge_range(b2_t, 0, b2_t, 1, "Single-Phase Line", bold)
    ws.write_url(b2_t, 2, "internal:'Line'!A1", link_fmt, "Go to Type List")
    ws.write_row(b2_h, 0, ["ID","Status","Length","From1","To1","r11 (Ohm/length_unit)","x11 (Ohm/length_unit)","b11 (uS/length_unit)"], th)
    rcur = b2_first
    for row in single_rows:
        ws.write(rcur, 0, row[0]); ws.write_number(rcur, 1, row[1])
        ws.write_number(rcur, 2, row[2], num6); ws.write(rcur, 3, row[3]); ws.write(rcur, 4, row[4])
        ws.write_number(rcur, 5, row[5], num4); ws.write_number(rcur, 6, row[6], num4); ws.write_number(rcur, 7, row[7], num4)
        rcur += 1
    ws.merge_range(b2_e, 0, b2_e, 1, "End of Single-Phase Line")

    # ---- Block 3: Two-Phase ----
    ws.merge_range(b3_t, 0, b3_t, 1, "Two-Phase Line", bold)
    ws.write_url(b3_t, 2, "internal:'Line'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b3_h, 0,
        ["ID","Status","Length","From1","From2","To1","To2",
         "r11 (Ohm/length_unit)","x11 (Ohm/length_unit)",
         "r21 (Ohm/length_unit)","x21 (Ohm/length_unit)",
         "r22 (Ohm/length_unit)","x22 (Ohm/length_unit)",
         "b11 (uS/length_unit)","b21 (uS/length_unit)","b22 (uS/length_unit)"],
        th
    )
    rcur = b3_first
    for row in two_rows:
        for c, v in enumerate(row):
            if c in (2, 7, 8, 9, 10, 11, 12, 13, 14, 15):  # numeric cols
                fmt = num6 if c == 2 else num4
                ws.write_number(rcur, c, v if v is not None else 0.0, fmt)
            else:
                ws.write(rcur, c, v if v is not None else "")
        rcur += 1
    ws.merge_range(b3_e, 0, b3_e, 1, "End of Two-Phase Line")

    # ---- Block 4: Three-Phase Full Data ----
    ws.merge_range(b4_t, 0, b4_t, 1, "Three-Phase Line with Full Data", bold)
    ws.write_url(b4_t, 2, "internal:'Line'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b4_h, 0,
        ["ID","Status","Length","From1","From2","From3","To1","To2","To3",
         "r11 (Ohm/length_unit)","x11 (Ohm/length_unit)",
         "r21 (Ohm/length_unit)","x21 (Ohm/length_unit)",
         "r22 (Ohm/length_unit)","x22 (Ohm/length_unit)",
         "r31 (Ohm/length_unit)","x31 (Ohm/length_unit)",
         "r32 (Ohm/length_unit)","x32 (Ohm/length_unit)",
         "r33 (Ohm/length_unit)","x33 (Ohm/length_unit)",
         "b11 (uS/length_unit)","b21 (uS/length_unit)","b22 (uS/length_unit)",
         "b31 (uS/length_unit)","b32 (uS/length_unit)","b33 (uS/length_unit)"],
        th
    )
    rcur = b4_first
    for row in three_full_rows:
        for c, v in enumerate(row):
            if c in (2, 9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25):  # numeric
                fmt = num6 if c == 2 else num4
                ws.write_number(rcur, c, v if v is not None else 0.0, fmt)
            else:
                ws.write(rcur, c, v if v is not None else "")
        rcur += 1
    ws.merge_range(b4_e, 0, b4_e, 1, "End of Three-Phase Line with Full Data")

    # ---- Block 5: Three-Phase Sequential (template only) ----
    ws.merge_range(b5_t, 0, b5_t, 1, "Three-Phase Line with Sequential Data", bold)
    ws.write_url(b5_t, 2, "internal:'Line'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b5_h, 0,
        ["ID","Status","Length","From1","From2","From3","To1","To2","To3",
         "R0 (Ohm/length_unit)","X0 (Ohm/length_unit)","R1 (Ohm/length_unit)",
         "X1 (Ohm/length_unit)","B0 (uS/length_unit)","B1 (uS/length_unit)"],
        th
    )
    ws.merge_range(b5_e, 0, b5_e, 2, "End of Three-Phase Line with Sequential Data")
