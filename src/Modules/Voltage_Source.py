# Modules/Voltage_Source.py
from __future__ import annotations
from pathlib import Path
import xml.etree.ElementTree as ET
from typing import List, Dict, Any, Optional
import re

# --- sanitizers (no '-' allowed) ---
_SAFE_RE = re.compile(r"[^A-Za-z0-9_]+")
def _safe_name(s: Optional[str]) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    s = _SAFE_RE.sub("_", s)
    s = re.sub(r"_+", "_", s)
    return s.strip("_")

def _make_src_id(raw: str) -> str:
    """
    Sanitize and ensure single 'SRC_' prefix.
    """
    name = _safe_name(raw)
    if not name:
        return "SRC_?"
    return name if name.startswith("SRC_") else f"SRC_{name}"


def _f(x: Optional[str]) -> Optional[float]:
    try:
        return float(x) if x not in (None, "") else None
    except Exception:
        return None


def _get_source_id(src: ET.Element, node: str, model_index: int | None = None) -> str:
    sid_raw = (
        (src.findtext("./SourceSettings/SourceID") or "").strip()
        or (src.findtext("./SourceSettings/DeviceNumber") or "").strip()
        or (src.findtext("./SourceID") or "").strip()
    )
    base = sid_raw if sid_raw else _safe_name(node)
    if model_index and model_index > 1:
        base = f"{base}_M{model_index}"   # underscore, not hyphen
    return _make_src_id(base)


def _pick_seq_impedances(eq: ET.Element | None, src: ET.Element) -> Optional[Dict[str, float]]:
    use_second = (
        (eq.findtext("UseSecondLevelImpedance") if eq is not None else None)
        or src.findtext("UseSecondLevelImpedance")
        or "0"
    )
    use_second = use_second.strip() == "1"

    def pick(name_first: str, name_second: str) -> Optional[float]:
        v2 = _f(eq.findtext(name_second)) if (use_second and eq is not None) else None
        if v2 is not None:
            return v2
        v1 = _f(eq.findtext(name_first)) if eq is not None else None
        return v1

    if eq is None:
        return None

    r1 = pick("FirstLevelR1", "SecondLevelR1")
    x1 = pick("FirstLevelX1", "SecondLevelX1")
    r0 = pick("FirstLevelR0", "SecondLevelR0")
    x0 = pick("FirstLevelX0", "SecondLevelX0")

    if None in (r1, x1, r0, x0):
        return None
    return {"R1": float(r1), "X1": float(x1), "R0": float(r0), "X0": float(x0)}


def _pick_sc_capacities(eq: ET.Element | None, src: ET.Element) -> Optional[Dict[str, float]]:
    vals: List[Optional[float]] = []
    if eq is not None:
        vals.extend([
            _f(eq.findtext("ShortCircuitMVA")),
            _f(eq.findtext("NominalCapacity1MVA")),
            _f(eq.findtext("NominalCapacity2MVA")),
        ])
    vals.append(_f(src.findtext("ShortCircuitMVA")))
    found = [v for v in vals if v and v > 0]
    if not found:
        return None
    if len(found) >= 2:
        return {"SC1ph": float(found[0]), "SC3ph": float(found[1])}
    v = float(found[0])
    return {"SC1ph": v, "SC3ph": v}


def _gather_sources(root: ET.Element) -> List[ET.Element]:
    """
    Prefer sources under Topo blocks whose NetworkType == 'Substation'.
    Ignore Topo blocks that are Feeders (and those with EquivalentMode == 1).
    If no Topo/Substation is found, fall back to scanning all .//Sources/Source.
    """
    picked: List[ET.Element] = []

    topo_nodes = root.findall(".//Topo")
    for topo in topo_nodes:
        ntype = (topo.findtext("NetworkType") or "").strip().lower()
        eq_mode = (topo.findtext("EquivalentMode") or "").strip()
        if ntype != "substation":
            continue
        if eq_mode == "1":
            continue
        srcs_parent = topo.find("./Sources")
        if srcs_parent is None:
            continue
        picked.extend(srcs_parent.findall("./Source"))

    if picked:
        return picked

    return root.findall(".//Sources/Source")


def _parse_voltage_sources(path: Path) -> tuple[list[dict], list[dict]]:
    """
    Returns (sc_rows, seq_rows):
      - sc_rows  -> Short-Circuit Level Data rows
      - seq_rows -> Sequential Data rows
    Only includes sources from Topo blocks where NetworkType == 'Substation'.
    """
    txt = path.read_text(encoding="utf-8", errors="ignore")
    root = ET.fromstring(txt)

    sc_rows: List[Dict[str, Any]] = []
    seq_rows: List[Dict[str, Any]] = []

    for src in _gather_sources(root):
        node_raw = (src.findtext("./SourceNodeID") or "").strip()
        if not node_raw:
            continue
        node = _safe_name(node_raw)  # sanitize bus base name

        models = src.findall("./EquivalentSourceModels/EquivalentSourceModel")
        if not models:
            models = [src]

        for idx, mdl in enumerate(models, start=1):
            eq = mdl.find("./EquivalentSource")

            kvll = (
                (eq.findtext("KVLL") if eq is not None else None)
                or mdl.findtext("KVLL")
                or src.findtext("KVLL")
            )
            kV = _f(kvll)
            if kV is None:
                continue

            # Angle default: 0.0 if not provided (no offset)
            angle_a = (
                _f(eq.findtext("OperatingAngle1")) if eq is not None else None
            ) or _f(mdl.findtext("OperatingAngle1")) or _f(src.findtext("OperatingAngle1"))
            if angle_a is None:
                angle_a = 0.0

            sid = _get_source_id(src, node, model_index=idx if len(models) > 1 else None)

            seq = _pick_seq_impedances(eq, src)
            if seq:
                seq_rows.append(
                    {
                        "ID": sid,
                        "Bus1": f"{node}_a",
                        "Bus2": f"{node}_b",
                        "Bus3": f"{node}_c",
                        "kV": kV,
                        "Angle_a": float(angle_a),
                        "R1": seq["R1"], "X1": seq["X1"], "R0": seq["R0"], "X0": seq["X0"],
                    }
                )
            else:
                caps = _pick_sc_capacities(eq, src)
                sc1 = 200000.0
                sc3 = 200000.0
                if caps:
                    sc1 = caps["SC1ph"]
                    sc3 = caps["SC3ph"]

                sc_rows.append(
                    {
                        "ID": sid,
                        "Bus1": f"{node}_a",
                        "Bus2": f"{node}_b",
                        "Bus3": f"{node}_c",
                        "kV": kV,
                        "Angle_a": float(angle_a),
                        "SC1ph": sc1,
                        "SC3ph": sc3,
                    }
                )

    return sc_rows, seq_rows


def write_voltage_source_sheet(xw, input_path: Path) -> None:
    wb = xw.book
    ws = wb.add_worksheet("Voltage Source")
    xw.sheets["Voltage Source"] = ws

    # Formats
    bold = wb.add_format({"bold": True})
    link_fmt = wb.add_format({"font_color": "blue", "underline": 1})
    notes_hdr = wb.add_format({"bold": True, "font_color": "yellow", "bg_color": "#595959", "align": "left"})
    notes_txt = wb.add_format({"font_color": "yellow", "bg_color": "#595959"})
    th = wb.add_format({"bold": True, "bottom": 1})
    num = wb.add_format({"num_format": "0.00"})
    num_int = wb.add_format({"num_format": "0"})

    # Column widths A..J
    widths = [28, 14, 14, 14, 18, 14, 16, 16, 12, 12]
    for c, w in enumerate(widths):
        ws.set_column(c, c, w)

    # Header
    ws.write(0, 0, "Type", bold)

    # Notes
    ws.merge_range(7, 0, 7, 7, "Important notes:", notes_hdr)
    ws.merge_range(8, 0, 8, 7, "Default order of blocks and columns after row 11 must not change", notes_txt)
    ws.merge_range(
        9, 0, 9, 7,
        "One empty row between End of each block and the next block is mandatory; otherwise, empty rows are NOT allowed",
        notes_txt,
    )

    # Block rows layout
    r = 10
    b1_title = r; b1_header = r + 1; b1_end = r + 2; r = b1_end + 2
    b2_title = r; b2_header = r + 1; b2_end = r + 2; r = b2_end + 2

    sc_rows, seq_rows = _parse_voltage_sources(Path(input_path))

    b3_title = r; b3_header = r + 1; b3_first = r + 2; b3_end = b3_first + len(sc_rows); r = b3_end + 2
    b4_title = r; b4_header = r + 1; b4_first = r + 2; b4_end = b4_first + len(seq_rows)

    # Top links
    ws.write_url(1, 0, f"internal:'Voltage Source'!A{b1_header+1}:F{b1_end+1}", link_fmt, "PositiveSeqVsource")
    ws.write_url(2, 0, f"internal:'Voltage Source'!A{b2_header+1}:F{b2_end+1}", link_fmt, "SinglePhaseVsource")
    ws.write_url(3, 0, f"internal:'Voltage Source'!A{b3_header+1}:H{max(b3_header+1,b3_end+1)}", link_fmt, "ThreePhaseShortCircuitVsource")
    ws.write_url(4, 0, f"internal:'Voltage Source'!A{b4_header+1}:J{max(b4_header+1,b4_end+1)}", link_fmt, "ThreePhaseSequentialVsource")

    # Block 1 (placeholder)
    ws.merge_range(b1_title, 0, b1_title, 3, "Positive-Sequence Voltage Source", bold)
    ws.write_url(b1_title, 4, "internal:'Voltage Source'!A1", link_fmt, "Go to Type List")
    ws.write_row(b1_header, 0, ["ID", "Bus", "Voltage (pu)", "Angle (deg)", "Rs (pu)", "Xs (pu)"], th)
    ws.merge_range(b1_end, 0, b1_end, 3, "End of Positive-Sequence Voltage Source")

    # Block 2 (placeholder)
    ws.merge_range(b2_title, 0, b2_title, 3, "Single-Phase Voltage Source", bold)
    ws.write_url(b2_title, 4, "internal:'Voltage Source'!A1", link_fmt, "Go to Type List")
    ws.write_row(b2_header, 0, ["ID", "Bus1", "Voltage (V)", "Angle (deg)", "Rs (Ohm)", "Xs (Ohm)"], th)
    ws.merge_range(b2_end, 0, b2_end, 3, "End of Single-Phase Voltage Source")

    # Block 3 (Short-Circuit Level)
    ws.merge_range(b3_title, 0, b3_title, 3, "Three-Phase Voltage Source with Short-Circuit Level Data", bold)
    ws.write_url(b3_title, 4, "internal:'Voltage Source'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b3_header, 0,
        ["ID", "Bus1", "Bus2", "Bus3", "kV (ph-ph RMS)", " Angle_a (deg)", "SC1ph (MVA)", "SC3ph (MVA)"],
        th,
    )
    rcur = b3_first
    for s in sc_rows:
        ws.write(rcur, 0, s["ID"])
        ws.write(rcur, 1, s["Bus1"])
        ws.write(rcur, 2, s["Bus2"])
        ws.write(rcur, 3, s["Bus3"])
        ws.write_number(rcur, 4, float(s["kV"]), num)
        ws.write_number(rcur, 5, float(s["Angle_a"]), num)
        ws.write_number(rcur, 6, float(s["SC1ph"]), num)
        ws.write_number(rcur, 7, float(s["SC3ph"]), num)
        rcur += 1
    ws.merge_range(b3_end, 0, b3_end, 3, "End of Three-Phase Voltage Source with Short-Circuit Level Data")

    # Block 4 (Sequential Data)
    ws.merge_range(b4_title, 0, b4_title, 3, "Three-Phase Voltage Source with Sequential Data", bold)
    ws.write_url(b4_title, 4, "internal:'Voltage Source'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b4_header, 0,
        ["ID", "Bus1", "Bus2", "Bus3", "kV (ph-ph RMS)", " Angle_a (deg)", "R1 (Ohm)", "X1 (Ohm)", "R0 (Ohm)", "X0 (Ohm)"],
        th,
    )
    rcur = b4_first
    for s in seq_rows:
        ws.write(rcur, 0, s["ID"])
        ws.write(rcur, 1, s["Bus1"])
        ws.write(rcur, 2, s["Bus2"])
        ws.write(rcur, 3, s["Bus3"])
        ws.write_number(rcur, 4, float(s["kV"]), num)
        ws.write_number(rcur, 5, float(s["Angle_a"]), num)
        ws.write_number(rcur, 6, float(s["R1"]), num)
        ws.write_number(rcur, 7, float(s["X1"]), num)
        ws.write_number(rcur, 8, float(s["R0"]), num)
        ws.write_number(rcur, 9, float(s["X0"]), num)
        rcur += 1
    ws.merge_range(b4_end, 0, b4_end, 3, "End of Three-Phase Voltage Source with Sequential Data")
