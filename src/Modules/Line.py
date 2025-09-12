# Modules/Line.py
from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple
import re
from Modules.Bus import extract_bus_data  # reuse Bus page logic (and comment filtering)
from Modules.IslandFilter import should_comment_branch, should_drop_branch, drop_mode_enabled
from Modules.General import safe_name


# ---- constants / small helpers ----
MI_PER_M = 0.000621371192
MI_PER_KM = 0.621371192  # multiply (per-km) to get per-mile
PHASE_SUFFIX = {"A": "_a", "B": "_b", "C": "_c"}

_PHASE_SUFFIX_RE = re.compile(r"_(a|b|c)$")

def _bus_base_set_from_bus_sheet(input_path: Path) -> set[str]:
    """
    Build the set of *active* base bus names that appear on the Bus sheet.
    We skip commented rows (those with Bus starting with '//') and strip the trailing _a/_b/_c.
    """
    bases: set[str] = set()
    for row in extract_bus_data(input_path):
        b = str(row.get("Bus", "")).strip()
        if not b or b.startswith("//"):
            continue  # commented-out bus is NOT considered active/known
        base = _PHASE_SUFFIX_RE.sub("", b)
        if base:
            bases.add(base)
    return bases


def _f(s: Optional[str]) -> float:
    """Safe float: XML text -> float, never None."""
    try:
        return float(s) if s not in (None, "") else 0.0
    except Exception:
        return 0.0


# ---- read line/cable databases (per-length) ----
def _read_line_db_map(root: ET.Element) -> Dict[str, Dict[str, float]]:
    """
    EquipmentID -> dict of per-length values (assumed per-km for R/X and ÂµS/km for B).
    Supports BOTH:
      - per-phase set: SelfResistance*, SelfReactance*, ShuntSusceptance*, Mutual*
      - sequence set:  PositiveSequenceResistance/Reactance/ShuntSusceptance,
                       ZeroSequenceResistance/Reactance/ShuntSusceptance
    """
    per_phase_keys = [
        "SelfResistanceA", "SelfResistanceB", "SelfResistanceC",
        "SelfReactanceA",  "SelfReactanceB",  "SelfReactanceC",
        "ShuntSusceptanceA", "ShuntSusceptanceB", "ShuntSusceptanceC",
        "MutualResistanceAB", "MutualResistanceBC", "MutualResistanceCA",
        "MutualReactanceAB",  "MutualReactanceBC",  "MutualReactanceCA",
        "MutualShuntSusceptanceAB", "MutualShuntSusceptanceBC", "MutualShuntSusceptanceCA",
    ]
    seq_keys = [
        "PositiveSequenceResistance", "PositiveSequenceReactance", "PositiveSequenceShuntSusceptance",
        "ZeroSequenceResistance",     "ZeroSequenceReactance",     "ZeroSequenceShuntSusceptance",
    ]

    db: Dict[str, Dict[str, float]] = {}
    for node in root.iter():
        tag = node.tag.split("}")[-1] if isinstance(node.tag, str) else ""
        if not tag.endswith("DB"):
            continue
        eid = (node.findtext("EquipmentID") or "").strip()
        if not eid:
            continue

        vals: Dict[str, float] = {}
        has_any = False

        # collect per-phase fields
        for k in per_phase_keys:
            v = _f(node.findtext(k))
            vals[k] = v
            has_any = has_any or (v != 0.0)

        # collect positive/zero sequence fields
        for k in seq_keys:
            v = _f(node.findtext(k))
            vals[k] = v
            has_any = has_any or (v != 0.0)

        if has_any:
            db[eid] = vals

    return db


# ---- parse section blocks ----
def _iter_lines(root: ET.Element):
    """
    Yield dicts describing each line/cable section.
      {type, id, from, to, phase, length_m, line_id}

    NOTE: from/to and id are sanitized to [A-Za-z0-9_].
    The database lookup key (line_id/cable_id) is NOT sanitized.
    """
    for sec in root.findall(".//Sections/Section"):
        from_bus_raw = (sec.findtext("FromNodeID") or "").strip()
        to_bus_raw   = (sec.findtext("ToNodeID") or "").strip()
        phase        = (sec.findtext("Phase") or "").strip().upper() or "ABC"

        # sanitize bus names for output/IDs
        from_bus = safe_name(from_bus_raw)
        to_bus   = safe_name(to_bus_raw)
        row_id   = safe_name(f"LN_{from_bus}_{to_bus}")

        # OverheadLineUnbalanced
        olu = sec.find(".//Devices/OverheadLineUnbalanced")
        if olu is not None:
            yield {
                "type": "OverheadLineUnbalanced",
                "id": row_id,
                "from": from_bus, "to": to_bus,
                "phase": phase,
                "length_m": _f(olu.findtext("Length")),
                "line_id": (olu.findtext("LineID") or "").strip(),  # keep raw for DB match
            }
            continue

        # OverheadByPhase (no explicit LineID)
        obp = sec.find(".//Devices/OverheadByPhase")
        if obp is not None:
            yield {
                "type": "OverheadByPhase",
                "id": row_id,
                "from": from_bus, "to": to_bus,
                "phase": phase,
                "length_m": _f(obp.findtext("Length")),
                "line_id": "",
            }
            continue

        # Balanced OverheadLine (has LineID)
        ol = sec.find(".//Devices/OverheadLine")
        if ol is not None:
            yield {
                "type": "OverheadLine",
                "id": row_id,
                "from": from_bus, "to": to_bus,
                "phase": phase,
                "length_m": _f(ol.findtext("Length")),
                "line_id": (ol.findtext("LineID") or "").strip(),  # keep raw for DB match
            }
            continue

        # Underground variants (include 'Cable' device)
        ug = (
            sec.find(".//Devices/Underground")
            or sec.find(".//Devices/UndergroundCable")
            or sec.find(".//Devices/UndergroundCableUnbalanced")
            or sec.find(".//Devices/UndergroundByPhase")
            or sec.find(".//Devices/Cable")
        )
        if ug is not None:
            yield {
                "type": "Underground",
                "id": row_id,
                "from": from_bus, "to": to_bus,
                "phase": phase,
                "length_m": _f(ug.findtext("Length")),
                "line_id": (ug.findtext("CableID") or ug.findtext("LineID") or "").strip(),  # keep raw
            }
            continue


# ---- helpers to compute R/X/B per mile from DB ----
def _scaled(v: float) -> float:
    """per-km â†’ per-mile (R/X in Î©; B in ÂµS)."""
    return v * MI_PER_KM


def _has_per_phase(dbvals: Dict[str, float]) -> bool:
    """
    Any non-zero per-phase field means per-phase data exists.
    (Avoids misclassifying as sequence-only when per-phase fields are present.)
    """
    keys = (
        "SelfResistanceA", "SelfReactanceA", "ShuntSusceptanceA",
        "SelfResistanceB", "SelfReactanceB", "ShuntSusceptanceB",
        "SelfResistanceC", "SelfReactanceC", "ShuntSusceptanceC",
        "MutualResistanceAB", "MutualReactanceAB", "MutualShuntSusceptanceAB",
        "MutualResistanceBC", "MutualReactanceBC", "MutualShuntSusceptanceBC",
        "MutualResistanceCA", "MutualReactanceCA", "MutualShuntSusceptanceCA",
    )
    return any(abs(dbvals.get(k, 0.0)) > 0.0 for k in keys)


def _per_phase_matrix_from_seq(dbvals: Dict[str, float]) -> Tuple[float, float, float, float, float, float]:
    """
    When only sequence data is present, build equivalent phase-domain
    per-length numbers for a fully transposed line:
      Zself = (Z0 + 2*Z1)/3,  Zmut = (Z0 - Z1)/3
      Bself = (B0 + 2*B1)/3,  Bmut = (B0 - B1)/3
    Returns (r_self, x_self, b_self, r_mut, x_mut, b_mut) per MILE.
    """
    R1 = _scaled(dbvals.get("PositiveSequenceResistance", 0.0))
    X1 = _scaled(dbvals.get("PositiveSequenceReactance", 0.0))
    R0 = _scaled(dbvals.get("ZeroSequenceResistance", 0.0))
    X0 = _scaled(dbvals.get("ZeroSequenceReactance", 0.0))

    B1 = _scaled(dbvals.get("PositiveSequenceShuntSusceptance", 0.0))
    B0 = _scaled(dbvals.get("ZeroSequenceShuntSusceptance", 0.0))

    r_self = (R0 + 2.0 * R1) / 3.0
    x_self = (X0 + 2.0 * X1) / 3.0
    r_mut  = (R0 - R1) / 3.0
    x_mut  = (X0 - X1) / 3.0

    b_self = (B0 + 2.0 * B1) / 3.0
    b_mut  = (B0 - B1) / 3.0

    return r_self, x_self, b_self, r_mut, x_mut, b_mut


def _series_shunt_for_pair(dbvals: Dict[str, float], p1: str, p2: Optional[str] = None):
    """
    Return (r11, x11, b11, r21, x21, b21, r22, x22, b22) per mile.
    Works with either per-phase DBs or sequence-only DBs.
    - If sequence-only: use transposed-line relations above.
    - For 2-Ï†, we set r22/x22/b22 = r11/x11/b11 and r21/x21/b21 = mutual.
    """
    if _has_per_phase(dbvals):
        # per-phase data available
        r11 = _scaled(dbvals.get(f"SelfResistance{p1}", 0.0))
        x11 = _scaled(dbvals.get(f"SelfReactance{p1}", 0.0))
        b11 = _scaled(dbvals.get(f"ShuntSusceptance{p1}", 0.0))

        if not p2:
            return r11, x11, b11, None, None, None, None, None, None

        r22 = _scaled(dbvals.get(f"SelfResistance{p2}", 0.0))
        x22 = _scaled(dbvals.get(f"SelfReactance{p2}", 0.0))
        b22 = _scaled(dbvals.get(f"ShuntSusceptance{p2}", 0.0))

        pair = "".join(sorted([p1, p2]))  # AB, AC, BC
        key = {"AB": "AB", "AC": "CA", "BC": "BC"}[pair]  # DB uses AB, BC, CA
        r21 = _scaled(dbvals.get(f"MutualResistance{key}", 0.0)) * 2.0
        x21 = _scaled(dbvals.get(f"MutualReactance{key}", 0.0)) * 2.0
        b21 = _scaled(dbvals.get(f"MutualShuntSusceptance{key}", 0.0)) * 2.0
        return r11, x11, b11, r21, x21, b21, r22, x22, b22

    # sequence-only â†’ build equivalents
    r_s, x_s, b_s, r_m, x_m, b_m = _per_phase_matrix_from_seq(dbvals)

    if not p2:
        return r_s, x_s, b_s, None, None, None, None, None, None

    # two-phase: equal self on each phase, mutual between them
    return r_s, x_s, b_s, r_m * 2.0, x_m * 2.0, b_m * 2.0, r_s, x_s, b_s


# ---- sheet writer ----
def write_line_sheet(xw, input_path: Path) -> None:
    root = ET.fromstring(Path(input_path).read_text(encoding="utf-8", errors="ignore"))
    dbmap = _read_line_db_map(root)

    # Build known (ACTIVE) bus set from the Bus sheet logic
    known_buses = _bus_base_set_from_bus_sheet(input_path)

    def _mark_unknown(bus_base: str) -> str:
        # bus_base is sanitized already by pipeline
        return bus_base if bus_base in known_buses else f"{bus_base}_unknown"

    single_rows: List[List[object]] = []
    two_rows: List[List[object]] = []
    three_full_rows: List[List[object]] = []

    for item in _iter_lines(root):
        phase = item["phase"]

        # base names (already sanitized by _iter_lines)
        from_base = item["from"]
        to_base   = item["to"]

        # island filter decision **before** we add _unknown
        comment_for_island = should_comment_branch(from_base, to_base)

        # now append _unknown for endpoints that donâ€™t exist on the Bus sheet
        from_bus = _mark_unknown(from_base)
        to_bus   = _mark_unknown(to_base)

        unknown_from = from_bus.endswith("_unknown")
        unknown_to   = to_bus.endswith("_unknown")
        # Drop entirely only if BOTH endpoints are inactive/missing in drop mode.
        # Keep boundary edges (one endpoint in a different island) so they appear in the sheet.
        if drop_mode_enabled() and comment_for_island and (unknown_from and unknown_to):
            continue

        # Comment only if BOTH endpoints are inactive/missing; otherwise keep visible for troubleshooting.
        id_out = ("//" if (comment_for_island and (unknown_from and unknown_to)) else "") + item["id"]
        length_mi = (item["length_m"] or 0.0) * MI_PER_M
        status = 1

        # --- pick DB values BEFORE computing rows ---
        dbid = item["line_id"]
        if dbid and dbid in dbmap:
            dbvals = dbmap[dbid]
        elif item["type"] == "OverheadByPhase" and "DEFAULT" in dbmap:
            # Prefer file's DEFAULT DB when OverheadByPhase has no LineID/CableID
            dbvals = dbmap["DEFAULT"]
        else:
            # Keep old heuristic for other files
            fallback = "LINE601" if phase == "ABC" else "LINE603"
            dbvals = dbmap.get(fallback, {})

        if phase in ("A", "B", "C"):
            p = phase
            r11, x11, b11, *_ = _series_shunt_for_pair(dbvals, p)
            single_rows.append([
                id_out, status, length_mi,
                f"{from_bus}{PHASE_SUFFIX[p]}",
                f"{to_bus}{PHASE_SUFFIX[p]}",
                r11, x11, b11
            ])

        elif phase in ("AB", "BC", "AC"):
            p1, p2 = phase[0], phase[1]
            r11, x11, b11, r21, x21, b21, r22, x22, b22 = _series_shunt_for_pair(dbvals, p1, p2)
            two_rows.append([
                id_out, status, length_mi,
                f"{from_bus}{PHASE_SUFFIX[p1]}", f"{from_bus}{PHASE_SUFFIX[p2]}",
                f"{to_bus}{PHASE_SUFFIX[p1]}",   f"{to_bus}{PHASE_SUFFIX[p2]}",
                r11, x11, r21, x21, r22, x22, b11, b21, b22
            ])

        else:  # ABC
            if _has_per_phase(dbvals):
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
            else:
                # sequence-only â†’ build full matrix from Z1/Z0 and B1/B0
                r_s, x_s, b_s, r_m, x_m, b_m = _per_phase_matrix_from_seq(dbvals)
                r11 = r22 = r33 = r_s; x11 = x22 = x33 = x_s
                r21 = r31 = r32 = r_m; x21 = x31 = x32 = x_m
                b11 = b22 = b33 = b_s; b21 = b31 = b32 = b_m

            three_full_rows.append([
                id_out, status, length_mi,
                f"{from_bus}_a", f"{from_bus}_b", f"{from_bus}_c",
                f"{to_bus}_a",   f"{to_bus}_b",   f"{to_bus}_c",
                r11, x11, r21, x21, r22, x22, r31, x31, r32, x32, r33, x33,
                b11, b21, b22, b31, b32, b33
            ])

    # ---- write sheet (xlsxwriter) ----
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

    # Column widths
    widths = [28, 8, 12, 12, 12, 13, 13, 13, 12, 12, 12, 12, 12, 12,
              12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12]
    for c, w in enumerate(widths):
        ws.set_column(c, c, w)

    # Type + top links
    ws.write(0, 0, "Type", bold)

    # Anchors
    r = 10
    b1_t, b1_h, b1_e = r, r + 1, r + 2; r = b1_e + 2
    b2_t, b2_h, b2_first = r, r + 1, r + 2; b2_e = b2_first + len(single_rows); r = b2_e + 2
    b3_t, b3_h, b3_first = r, r + 1, r + 2; b3_e = b3_first + len(two_rows);  r = b3_e + 2
    b4_t, b4_h, b4_first = r, r + 1, r + 2; b4_e = b4_first + len(three_full_rows); r = b4_e + 2
    b5_t, b5_h, b5_e = r, r + 1, r + 2

    # Top link ranges
    ws.write_url(1, 0, f"internal:'Line'!A{b1_h+1}:G{b1_e+1}", link_fmt, "PositiveSeqLine")
    ws.write_url(2, 0, f"internal:'Line'!A{b2_h+1}:H{b2_e+1}", link_fmt, "SinglePhaseLine")
    ws.write_url(3, 0, f"internal:'Line'!A{b3_h+1}:P{b3_e+1}", link_fmt, "TwoPhaseLine")
    ws.write_url(4, 0, f"internal:'Line'!A{b4_h+1}:AA{max(b4_h+1, b4_e+1)}", link_fmt, "ThreePhaseLineFullData")
    ws.write_url(5, 0, f"internal:'Line'!A{b5_h+1}:O{b5_e+1}", link_fmt, "ThreePhaseLineSequentialData")

    # Notes
    ws.merge_range(7, 0, 7, 7, "Important notes:", notes_hdr)
    ws.merge_range(8, 0, 8, 7, "Default order of blocks and columns after row 11 must not change", notes_txt)
    ws.merge_range(9, 0, 9, 7, "One empty row between End of each block and the next block is mandatory; otherwise, empty rows are NOT allowed", notes_txt)

    # Block 1 (empty)
    ws.merge_range(b1_t, 0, b1_t, 1, "Positive-Sequence Line", bold)
    ws.write_url(b1_t, 2, "internal:'Line'!A1", link_fmt, "Go to Type List")
    ws.write_row(b1_h, 0, ["ID", "Status", "From bus", "To bus", "R (pu)", "X (pu)", "B (pu)"], th)
    ws.merge_range(b1_e, 0, b1_e, 1, "End of Positive-Sequence Line")

    # Block 2: Single-Phase
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

    # Block 3: Two-Phase
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

    # Block 4: Three-Phase Full Data
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

    # Block 5: Sequential (template)
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

