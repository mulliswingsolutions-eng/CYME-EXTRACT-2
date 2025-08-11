# Modules/Voltage_Source.py
from __future__ import annotations
from pathlib import Path
import xml.etree.ElementTree as ET
from typing import List, Dict, Any


def _parse_short_circuit_sources(path: Path) -> List[Dict[str, Any]]:
    txt = path.read_text(encoding="utf-8", errors="ignore")
    root = ET.fromstring(txt)

    out: List[Dict[str, Any]] = []
    sources = root.find(".//Sources")
    if sources is None:
        return out

    for src in sources.findall(".//Source"):
        node = (src.findtext(".//SourceNodeID") or "").strip()
        sid = (src.findtext(".//SourceID") or "").strip()
        kvll = (src.findtext(".//KVLL") or "").strip()
        ang = (
            src.findtext(".//EquivalentSource/OperatingAngle1")
            or src.findtext(".//OperatingAngle1")
            or ""
        ).strip()

        if not (node and sid and kvll and ang):
            continue

        try:
            kv = float(kvll)
        except ValueError:
            kv = None
        try:
            angle = float(ang)
        except ValueError:
            angle = None

        out.append({
            "ID": sid,
            "Bus1": f"{node}_a",
            "Bus2": f"{node}_b",
            "Bus3": f"{node}_c",
            "kV": kv,
            "Angle_a": angle,
            "SC1ph": 200000,  # hardcoded
            "SC3ph": 200000,  # hardcoded
        })
    return out


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

    # ---------- Type list (Rows 1–5) ----------
    ws.write(0, 0, "Type", bold)

    # Blocks start at row 11 (1-based) -> index 10
    r = 10

    # Block 1 rows
    b1_title = r
    b1_header = r + 1
    b1_end = r + 2
    r = b1_end + 2  # blank row between blocks

    # Block 2 rows
    b2_title = r
    b2_header = r + 1
    b2_end = r + 2
    r = b2_end + 2

    # Block 3 sized by parsed data
    data = _parse_short_circuit_sources(input_path)
    b3_title = r
    b3_header = r + 1
    b3_first = r + 2
    b3_end = b3_first + max(len(data), 0)
    r = b3_end + 2

    # Block 4 rows
    b4_title = r
    b4_header = r + 1
    b4_end = r + 2

    # ---------- Top links (select full table ranges) ----------
    # Block 1 table is A:F (6 columns)
    ws.write_url(1, 0, f"internal:'Voltage Source'!A{b1_header+1}:F{b1_end+1}", link_fmt, "PositiveSeqVsource")
    # Block 2 table is A:F
    ws.write_url(2, 0, f"internal:'Voltage Source'!A{b2_header+1}:F{b2_end+1}", link_fmt, "SinglePhaseVsource")
    # Block 3 table is A:H
    ws.write_url(3, 0, f"internal:'Voltage Source'!A{b3_header+1}:H{b3_end+1}", link_fmt, "ThreePhaseShortCircuitVsource")
    # Block 4 table is A:J (10 columns)
    ws.write_url(4, 0, f"internal:'Voltage Source'!A{b4_header+1}:J{b4_end+1}", link_fmt, "ThreePhaseSequentialVsource")

    # ---------- Important notes (start at row 8, merged A:H) ----------
    ws.merge_range(7, 0, 7, 7, "Important notes:", notes_hdr)  # Row 8
    ws.merge_range(8, 0, 8, 7, "Default order of blocks and columns after row 11 must not change", notes_txt)  # Row 9
    ws.merge_range(
        9, 0, 9, 7,
        "One empty row between End of each block and the next block is mandatory; otherwise, empty rows are NOT allowed",
        notes_txt,
    )  # Row 10

    # ---------- Block 1 ----------
    ws.merge_range(b1_title, 0, b1_title, 3, "Positive-Sequence Voltage Source", bold)  # A:D
    ws.write_url(b1_title, 4, "internal:'Voltage Source'!A1", link_fmt, "Go to Type List")  # Column E
    ws.write_row(b1_header, 0, ["ID", "Bus", "Voltage (pu)", "Angle (deg)", "Rs (pu)", "Xs (pu)"], th)
    ws.merge_range(b1_end, 0, b1_end, 3, "End of Positive-Sequence Voltage Source")  # A:D

    # ---------- Block 2 ----------
    ws.merge_range(b2_title, 0, b2_title, 3, "Single-Phase Voltage Source", bold)
    ws.write_url(b2_title, 4, "internal:'Voltage Source'!A1", link_fmt, "Go to Type List")
    ws.write_row(b2_header, 0, ["ID", "Bus1", "Voltage (V)", "Angle (deg)", "Rs (Ohm)", "Xs (Ohm)"], th)
    ws.merge_range(b2_end, 0, b2_end, 3, "End of Single-Phase Voltage Source")

    # ---------- Block 3 ----------
    ws.merge_range(b3_title, 0, b3_title, 3, "Three-Phase Voltage Source with Short-Circuit Level Data", bold)
    ws.write_url(b3_title, 4, "internal:'Voltage Source'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b3_header, 0,
        ["ID", "Bus1", "Bus2", "Bus3", "kV (ph-ph RMS)", "Angle_a (deg)", "SC1ph (MVA)", "SC3ph (MVA)"],
        th,
    )
    row = b3_first
    for s in data:
        ws.write(row, 0, s["ID"])
        ws.write(row, 1, s["Bus1"])
        ws.write(row, 2, s["Bus2"])
        ws.write(row, 3, s["Bus3"])
        ws.write_number(row, 4, s["kV"], num) if s["kV"] is not None else ws.write(row, 4, "")
        ws.write_number(row, 5, s["Angle_a"], num) if s["Angle_a"] is not None else ws.write(row, 5, "")
        ws.write_number(row, 6, s["SC1ph"], num_int)
        ws.write_number(row, 7, s["SC3ph"], num_int)
        row += 1
    ws.merge_range(b3_end, 0, b3_end, 3, "End of Three-Phase Voltage Source with Short-Circuit Level Data")

    # ---------- Block 4 ----------
    ws.merge_range(b4_title, 0, b4_title, 3, "Three-Phase Voltage Source with Sequential Data", bold)
    ws.write_url(b4_title, 4, "internal:'Voltage Source'!A1", link_fmt, "Go to Type List")
    ws.write_row(
        b4_header, 0,
        ["ID", "Bus1", "Bus2", "Bus3", "kV (ph-ph RMS)", "Angle_a (deg)", "R1 (Ohm)", "X1 (Ohm)", "R0 (Ohm)", "X0 (Ohm)"],
        th,
    )
    ws.merge_range(b4_end, 0, b4_end, 3, "End of Three-Phase Voltage Source with Sequential Data")
