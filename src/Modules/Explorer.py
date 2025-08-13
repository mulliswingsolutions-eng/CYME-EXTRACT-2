# src/Modules/Explorer.py
from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET
from typing import Any, Dict, Iterable, List, Optional, Tuple


def _txt(node: ET.Element | None, tag: str, default: str = "") -> str:
    if node is None:
        return default
    t = node.findtext(tag)
    return t.strip() if t else default


def _f(s: Optional[str]) -> Optional[float]:
    try:
        return float(s) if s not in (None, "") else None
    except Exception:
        return None


def _phase_list(ph: str) -> List[str]:
    ph = (ph or "").upper().strip()
    out: List[str] = []
    for ch in ("A", "B", "C"):
        if ch in ph:
            out.append(ch)
    return out


def _iter_sections(root: ET.Element) -> Iterable[Tuple[ET.Element, str, str, str, ET.Element | None, str]]:
    """
    Yields: (section_element, section_id, from_node, to_node, primary_device_element, primary_device_tag)
    We treat the *first* device child as the "primary" for classification.
    """
    for sec in root.findall(".//Sections/Section"):
        sec_id = _txt(sec, "SectionID")
        from_node = _txt(sec, "FromNodeID")
        to_node = _txt(sec, "ToNodeID")
        phase = _txt(sec, "Phase").upper() or ""
        devs = sec.find("Devices")
        primary = None
        ptag = ""
        if devs is not None:
            for child in list(devs):
                # pick the first element child as primary device
                if isinstance(child.tag, str):
                    primary = child
                    ptag = child.tag
                    break
        yield (sec, sec_id, from_node, to_node, primary, ptag)


# ------------------------- Extractors per device kind -------------------------

def _extract_sources(root: ET.Element) -> List[List[Any]]:
    """
    Returns rows:
      ["SourceNodeID","KVLL (kV)","V1 (kV LN)","V2 (kV LN)","V3 (kV LN)","Angle1 (deg)","Angle2 (deg)","Angle3 (deg)","Config"]
    """
    out: List[List[Any]] = []
    for src in root.findall(".//Sources/Source"):
        node = _txt(src, "SourceNodeID")
        cfg = _txt(src, "EquivalentSourceConfiguration")
        # EquivalentSource values may be under nested block
        eq = src.find(".//EquivalentSource")
        kvll = _f(_txt(eq, "KVLL")) if eq is not None else _f(_txt(src, "KVLL"))
        v1 = _f(_txt(eq, "OperatingVoltage1")) if eq is not None else _f(_txt(src, "OperatingVoltage1"))
        v2 = _f(_txt(eq, "OperatingVoltage2")) if eq is not None else _f(_txt(src, "OperatingVoltage2"))
        v3 = _f(_txt(eq, "OperatingVoltage3")) if eq is not None else _f(_txt(src, "OperatingVoltage3"))
        a1 = _f(_txt(eq, "OperatingAngle1")) if eq is not None else _f(_txt(src, "OperatingAngle1"))
        a2 = _f(_txt(eq, "OperatingAngle2")) if eq is not None else _f(_txt(src, "OperatingAngle2"))
        a3 = _f(_txt(eq, "OperatingAngle3")) if eq is not None else _f(_txt(src, "OperatingAngle3"))
        out.append([node, kvll, v1, v2, v3, a1, a2, a3, cfg])
    return out


def _extract_overhead_by_phase(sec_el: ET.Element, sec_id: str, f: str, t: str, phase: str, primary: ET.Element) -> List[Any]:
    return [
        sec_id, f, t, phase,
        _f(_txt(primary, "Length")),
        _txt(primary, "ConductorSpacingID"),
        _txt(primary, "PhaseConductorIDA"),
        _txt(primary, "PhaseConductorIDB"),
        _txt(primary, "PhaseConductorIDC"),
        _txt(primary, "NeutralConductorID1"),
        _txt(primary, "NeutralConductorID2"),
    ]


def _extract_overhead_unbalanced(sec_el: ET.Element, sec_id: str, f: str, t: str, phase: str, primary: ET.Element) -> List[Any]:
    return [
        sec_id, f, t, phase,
        _f(_txt(primary, "Length")),
        _txt(primary, "LineID"),
    ]


def _extract_underground(sec_el: ET.Element, sec_id: str, f: str, t: str, phase: str, primary: ET.Element) -> List[Any]:
    return [
        sec_id, f, t, phase,
        _f(_txt(primary, "Length")),
        _txt(primary, "CableID"),
        _txt(primary, "CableConfiguration"),
        _txt(primary, "BondingType"),
        _f(_txt(primary, "NumberOfCableInParallel")),
    ]


def _extract_switch_like(sec_el: ET.Element, sec_id: str, f: str, t: str, phase: str, primary: ET.Element) -> List[Any]:
    # works for Switch, Fuse, Breaker
    return [
        sec_id, f, t, phase,
        _txt(primary, "DeviceNumber"),
        _txt(primary, "NormalStatus"),
        _txt(primary, "ClosedPhase"),
        _txt(primary, "RemoteControlled"),  # empty for non-breakers
    ]


def _extract_transformer(sec_el: ET.Element, sec_id: str, f: str, t: str, phase: str, primary: ET.Element) -> List[Any]:
    return [
        sec_id, f, t, phase,
        _txt(primary, "DeviceID") or _txt(primary, "DeviceNumber"),
        _txt(primary, "TransformerConnection"),
        _txt(primary, "PhaseShift"),
        _f(_txt(primary, "PrimaryTapSettingPercent")),
        _f(_txt(primary, "SecondaryTapSettingPercent")),
    ]


def _sum_load_kw_kvar(load_parent: ET.Element) -> Tuple[Dict[str, float], Dict[str, float]]:
    """
    Sums per-phase KW/KVAR within CustomerLoadValues.
    Returns (kw_by_phase, kvar_by_phase), with total in keys "TOTAL".
    """
    kw: Dict[str, float] = {"A": 0.0, "B": 0.0, "C": 0.0, "TOTAL": 0.0}
    kvar: Dict[str, float] = {"A": 0.0, "B": 0.0, "C": 0.0, "TOTAL": 0.0}
    for cv in load_parent.findall(".//CustomerLoadValues/CustomerLoadValue"):
        ph = _txt(cv, "Phase").upper() or "TOTAL"
        # KW/KVAR appear inside <LoadValue ...><KW>,<KVAR>
        lv = cv.find("LoadValue")
        kw_v = _f(_txt(lv, "KW")) if lv is not None else None
        kvar_v = _f(_txt(lv, "KVAR")) if lv is not None else None
        if kw_v is not None:
            kw[ph if ph in "ABC" else "TOTAL"] += kw_v
            kw["TOTAL"] += kw_v
        if kvar_v is not None:
            kvar[ph if ph in "ABC" else "TOTAL"] += kvar_v
            kvar["TOTAL"] += kvar_v
    return kw, kvar


def _extract_spot_or_dist_load(sec_el: ET.Element, sec_id: str, f: str, t: str, phase: str, primary: ET.Element) -> List[Any]:
    conn = _txt(primary, "ConnectionConfiguration")
    kw, kvar = _sum_load_kw_kvar(primary)
    return [
        sec_id, f, phase or "", conn or "",
        kw["A"] or 0.0, kvar["A"] or 0.0,
        kw["B"] or 0.0, kvar["B"] or 0.0,
        kw["C"] or 0.0, kvar["C"] or 0.0,
        kw["TOTAL"] or 0.0, kvar["TOTAL"] or 0.0,
    ]


def _extract_shunt_cap(sec_el: ET.Element, sec_id: str, f: str, t: str, phase: str, primary: ET.Element) -> List[Any]:
    return [
        sec_id, f, phase or "",
        _f(_txt(primary, "KVLN")),
        _f(_txt(primary, "FixedKVARA")),
        _f(_txt(primary, "FixedKVARB")),
        _f(_txt(primary, "FixedKVARC")),
        _txt(primary, "ConnectionConfiguration"),
    ]


# ------------------------- Public writer -------------------------

def write_overview_sheet(xw, input_path: Path) -> None:
    """
    Writes an 'Overview' sheet that organizes the CYME export into readable blocks with hyperlinks.
    Index space is reserved *before* blocks to avoid any overlap.
    """
    text = Path(input_path).read_text(encoding="utf-8", errors="ignore")
    root = ET.fromstring(text)

    wb = xw.book
    ws = wb.add_worksheet("Overview")
    xw.sheets["Overview"] = ws

    # Formatting
    bold = wb.add_format({"bold": True})
    th = wb.add_format({"bold": True, "bottom": 1})
    link = wb.add_format({"font_color": "blue", "underline": 1})
    num = wb.add_format({"num_format": "0.000"})
    num0 = wb.add_format({"num_format": "0"})

    # Widths
    for col, w in enumerate([22, 18, 18, 10, 12, 12, 12, 12, 12, 12, 16, 16, 16, 16, 16]):
        ws.set_column(col, col, w)

    # ---- Build datasets ----
    sources_rows = _extract_sources(root)

    obp_rows: List[List[Any]] = []
    olu_rows: List[List[Any]] = []
    ug_rows: List[List[Any]] = []
    sw_rows: List[List[Any]] = []    # Switch/Fuse/Breaker
    xf_rows: List[List[Any]] = []
    sh_rows: List[List[Any]] = []
    sl_rows: List[List[Any]] = []    # SpotLoad / DistributedLoad

    for sec, sid, f, t, primary, ptag in _iter_sections(root):
        phase = _txt(sec, "Phase").upper() or ""
        if primary is None:
            continue

        if ptag == "OverheadByPhase":
            obp_rows.append(_extract_overhead_by_phase(sec, sid, f, t, phase, primary))
        elif ptag == "OverheadLineUnbalanced":
            olu_rows.append(_extract_overhead_unbalanced(sec, sid, f, t, phase, primary))
        elif ptag == "Underground":
            ug_rows.append(_extract_underground(sec, sid, f, t, phase, primary))
        elif ptag in ("Switch", "Fuse", "Breaker"):
            sw_rows.append(_extract_switch_like(sec, sid, f, t, phase, primary))
        elif ptag == "Transformer":
            xf_rows.append(_extract_transformer(sec, sid, f, t, phase, primary))
        elif ptag == "ShuntCapacitor":
            sh_rows.append(_extract_shunt_cap(sec, sid, f, t, phase, primary))
        elif ptag in ("SpotLoad", "DistributedLoad"):
            sl_rows.append(_extract_spot_or_dist_load(sec, sid, f, t, phase, primary))
        else:
            pass

    # ---- Plan blocks & reserve index space at the top ----
    planned_blocks: List[Tuple[str, List[List[Any]]]] = [
        ("Sources", sources_rows),
        ("Overhead Lines (ByPhase)", obp_rows),
        ("Overhead Lines (UnbalancedDB)", olu_rows),
        ("Underground Cables", ug_rows),
        ("Transformers", xf_rows),
        ("Switchgear (Switch / Fuse / Breaker)", sw_rows),
        ("Shunt Capacitors", sh_rows),
        ("Loads (Spot & Distributed)", sl_rows),
    ]

    # Index layout: header at row 0, links start at row 2
    index_header_row = 0
    index_first_link_row = 2
    index_rows = len(planned_blocks)
    reserved_rows_for_index = index_first_link_row + index_rows  # last link row index
    # First block starts *after* reserved index + one blank row
    r = reserved_rows_for_index + 1

    # Write index header now
    ws.write(index_header_row, 0, "Index", bold)

    # Helper to write a typed block with title, headers, and rows
    def write_block(title: str, headers: List[str], rows: List[List[Any]], number_cols: set[int] | None = None) -> Tuple[int, int, int]:
        """
        Returns (title_row, header_row, end_row_exclusive)
        """
        nonlocal r
        title_row = r
        ws.write(title_row, 0, title, bold)
        header_row = title_row + 1
        ws.write_row(header_row, 0, headers, th)
        body_row = header_row + 2  # start of data
        rr = body_row
        number_cols = number_cols or set()

        for row in rows:
            for c, v in enumerate(row):
                if isinstance(v, (int, float)) and c in number_cols:
                    ws.write_number(rr, c, v, num)
                elif isinstance(v, (int, float)):
                    ws.write_number(rr, c, v)
                else:
                    ws.write(rr, c, v)
            rr += 1

        # If no rows, keep an empty line to keep anchors valid
        if not rows:
            rr = body_row

        r = rr + 1  # blank line after block
        return (title_row, header_row, rr)

    # Write blocks and remember their title row for links
    anchors: List[Tuple[str, int]] = []

    # Sources
    trow, _, _ = write_block(
        "Sources",
        ["SourceNodeID","KVLL (kV)","V1 (kV LN)","V2 (kV LN)","V3 (kV LN)","Angle1 (deg)","Angle2 (deg)","Angle3 (deg)","Config"],
        sources_rows,
        number_cols={1,2,3,4,5,6,7}
    )
    anchors.append(("Sources", trow))

    # Overhead lines by phase
    trow, _, _ = write_block(
        "Overhead Lines (ByPhase)",
        ["SectionID","From","To","Phase","Length (m)","SpacingID","ConductorA","ConductorB","ConductorC","Neutral1","Neutral2"],
        obp_rows,
        number_cols={4}
    )
    anchors.append(("Overhead Lines (ByPhase)", trow))

    # Overhead lines (unbalanced DB)
    trow, _, _ = write_block(
        "Overhead Lines (UnbalancedDB)",
        ["SectionID","From","To","Phase","Length (m)","LineID"],
        olu_rows,
        number_cols={4}
    )
    anchors.append(("Overhead Lines (UnbalancedDB)", trow))

    # Underground cables
    trow, _, _ = write_block(
        "Underground Cables",
        ["SectionID","From","To","Phase","Length (m)","CableID","Configuration","Bonding","Parallel"],
        ug_rows,
        number_cols={4,8}
    )
    anchors.append(("Underground Cables", trow))

    # Transformers
    trow, _, _ = write_block(
        "Transformers",
        ["SectionID","From","To","Phase","DeviceID","Connection","PhaseShift","PrimaryTap (%)","SecondaryTap (%)"],
        xf_rows,
        number_cols={7,8}
    )
    anchors.append(("Transformers", trow))

    # Switchgear (Switch/Fuse/Breaker)
    trow, _, _ = write_block(
        "Switchgear (Switch / Fuse / Breaker)",
        ["SectionID","From","To","Phase","DeviceNumber","NormalStatus","ClosedPhase","RemoteControlled"],
        sw_rows,
    )
    anchors.append(("Switchgear (Switch / Fuse / Breaker)", trow))

    # Shunt capacitors
    trow, _, _ = write_block(
        "Shunt Capacitors",
        ["SectionID","Bus(From)","Phase","KVLN (kV LN)","FixedKVAR_A","FixedKVAR_B","FixedKVAR_C","Conn."],
        sh_rows,
        number_cols={3,4,5,6}
    )
    anchors.append(("Shunt Capacitors", trow))

    # Loads (Spot & Distributed)
    trow, _, _ = write_block(
        "Loads (Spot & Distributed)",
        ["SectionID","Bus(From)","Phase","Conn","KW_A","KVAR_A","KW_B","KVAR_B","KW_C","KVAR_C","KW_Total","KVAR_Total"],
        sl_rows,
        number_cols={4,5,6,7,8,9,10,11}
    )
    anchors.append(("Loads (Spot & Distributed)", trow))

    # ---- Now populate the reserved index area with working links (no overlap) ----
    current = index_first_link_row
    for label, rows_data in planned_blocks:
        # find anchor row for this label
        anchor_row = next((ar for (lab, ar) in anchors if lab == label), None)
        count = len(rows_data)
        # If no block written (shouldn't happen), link to A1
        target_row = (anchor_row + 1) if anchor_row is not None else 1
        ws.write_url(current, 0, f"internal:'Overview'!A{target_row}", link, f"{label}  (n={count})")
        current += 1
