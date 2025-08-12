# Modules/Transformer.py
from __future__ import annotations
from pathlib import Path
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple, Any


def _phase_count(phase_str: str) -> int:
    s = (phase_str or "").upper()
    return len({p for p in s if p in "ABC"}) or 3


def _conn_text(conn: str) -> str:
    c = (conn or "").upper()
    if "Y" in c:
        return "wye"
    if "D" in c:
        return "delta"
    return "wye"


def _bus_labels(bus: str, phase_str: str) -> Tuple[str, str, str]:
    s = (phase_str or "ABC").upper()
    labs = []
    for p in "ABC":
        labs.append(f"{bus}_{p.lower()}" if p in s else "")
    return tuple(labs)


def _compute_x_rw(z_percent: float, xr_ratio: float) -> Tuple[float, float, float]:
    """
    From PositiveSequenceImpedancePercent (Z%) and XRRatio, compute:
      - X (pu)
      - RW1 (pu) and RW2 (pu): split total R equally
    """
    try:
        z_pu = float(z_percent) / 100.0
        xr = float(xr_ratio)
    except Exception:
        return 0.0, 0.0, 0.0

    if z_pu <= 0 or xr <= 0:
        return 0.0, 0.0, 0.0

    denom = (1.0 + xr * xr) ** 0.5
    r_total = z_pu / denom
    x_pu = z_pu * xr / denom
    rw = r_total / 2.0
    return round(x_pu, 5), round(rw, 8), round(rw, 8)


def _read_transformer_db(root: ET.Element) -> Dict[str, Dict[str, Any]]:
    """
    TransformerDB block: key by EquipmentID.
    Fields used:
      PrimaryVoltage, SecondaryVoltage, NominalRatingKVA,
      PositiveSequenceImpedancePercent, XRRatio, TransformerConnection
    """
    out: Dict[str, Dict[str, Any]] = {}
    for tdb in root.findall(".//TransformerDB"):
        eid = (tdb.findtext("EquipmentID") or "").strip()
        if not eid:
            continue

        def f(name: str) -> str:
            v = tdb.findtext(name)
            return "" if v is None else v.strip()

        out[eid] = {
            "kvp": float(f("PrimaryVoltage") or 0) or float(f("PrimaryKV") or 0),
            "kvs": float(f("SecondaryVoltage") or 0) or float(f("SecondaryKV") or 0),
            "kva": float(f("NominalRatingKVA") or 0) or float(f("NominalRating") or 0),
            "z_pct": float(f("PositiveSequenceImpedancePercent") or 0),
            "xr": float(f("XRRatio") or 0),
            "conn": f("TransformerConnection") or f("Connection") or "Yg",
        }
    return out


def _parse_multiphase_2w_rows(input_path: Path) -> List[List[Any]]:
    """
    Parse <Section><Devices><Transformer> entries and build the Multiphase 2W rows.
    No synthetic rows added—only what exists in the file.
    """
    root = ET.fromstring(input_path.read_text(encoding="utf-8", errors="ignore"))
    tdb = _read_transformer_db(root)

    rows: List[List[Any]] = []

    for sec in root.findall(".//Section"):
        xf = sec.find(".//Devices/Transformer")
        if xf is None:
            continue

        from_bus = (sec.findtext("FromNodeID") or "").strip()
        to_bus = (sec.findtext("ToNodeID") or "").strip()
        phase = (sec.findtext("Phase") or "ABC").strip().upper()
        status_text = (xf.findtext("ConnectionStatus") or "Connected").strip().lower()
        status = 1 if status_text == "connected" else 0
        dev_id = (xf.findtext("DeviceID") or "").strip()

        # Lookup transformer DB
        info = tdb.get(dev_id, {})
        kvp = float(info.get("kvp") or 0.0)
        kvs = float(info.get("kvs") or 0.0)
        kva = float(info.get("kva") or 0.0)
        conn = _conn_text(info.get("conn", ""))
        x_pu, rw1, rw2 = _compute_x_rw(info.get("z_pct", 0.0), info.get("xr", 0.0))

        # Labels with phases
        b1a, b1b, b1c = _bus_labels(from_bus, phase)
        b2a, b2b, b2c = _bus_labels(to_bus, phase)

        rid = f"TR1_{from_bus}_{to_bus}"
        rows.append([
            rid, status, _phase_count(phase),
            b1a, b1b, b1c, kvp, kva, conn,
            b2a, b2b, b2c, kvs, kva, conn,
            0, 0, 0, -16, 16, 10, 10,
            x_pu, rw1, rw2
        ])

    return rows


def write_transformer_sheet(xw, input_path: Path) -> None:
    """
    Build the 'Transformer' sheet with the same layout rules as other pages.
    - Notes start at row 8 (merged A:H)
    - Blocks start on row 11
    - Titles merged A:D, 'Go to Type List' in column E (links to A1)
    - One blank row between blocks
    - Top 'Type' links select the full table ranges:
        PosSeq2wXF: A:K
        PosSeq3wXF: A:S
        Multiphase2wXF: A:Y
        Multiphase2wXFMutual: A:AA
    """
    wb = xw.book
    ws = wb.add_worksheet("Transformer")
    xw.sheets["Transformer"] = ws

    # Formats
    bold = wb.add_format({"bold": True})
    link_fmt = wb.add_format({"font_color": "blue", "underline": 1})
    notes_hdr = wb.add_format({"bold": True, "font_color": "yellow", "bg_color": "#595959", "align": "left"})
    notes_txt = wb.add_format({"font_color": "yellow", "bg_color": "#595959"})
    th = wb.add_format({"bold": True, "bottom": 1})
    num2 = wb.add_format({"num_format": "0.00"})
    num5 = wb.add_format({"num_format": "0.00000"})
    num8 = wb.add_format({"num_format": "0.00000000"})
    int0 = wb.add_format({"num_format": "0"})

    # Column widths
    widths = [22, 8, 16, 12, 12, 12, 8, 12, 10, 12, 12, 12, 8, 12, 10,
              8, 8, 8, 12, 12, 14, 14, 10, 12, 12, 12, 12, 12, 14, 14]
    for c, w in enumerate(widths):
        ws.set_column(c, c, w)

    # Precompute data for MP 2W block
    rows_mp2w = _parse_multiphase_2w_rows(input_path)

    # Blocks start at row 11 (1-based) => index 10
    r = 10

    # Block 1: Positive-Sequence 2W (empty)
    b1_title = r; b1_hdr = r + 1; b1_end = r + 2; r = b1_end + 2
    # Block 2: Positive-Sequence 3W (empty)
    b2_title = r; b2_hdr = r + 1; b2_end = r + 2; r = b2_end + 2
    # Block 3: Multiphase 2W (data)
    b3_title = r; b3_hdr = r + 1; b3_first = r + 2
    b3_end = b3_first + len(rows_mp2w); r = b3_end + 2
    # Block 4: Multiphase 2W w/ Mutual (empty)
    b4_title = r; b4_hdr = r + 1; b4_end = r + 2

    # Type row + links
    ws.write(0, 0, "Type", bold)
    ws.write_url(1, 0, f"internal:'Transformer'!A{b1_hdr+1}:K{b1_end+1}", link_fmt, "PositiveSeq2wXF")
    ws.write_url(2, 0, f"internal:'Transformer'!A{b2_hdr+1}:S{b2_end+1}", link_fmt, "PositiveSeq3wXF")
    ws.write_url(3, 0, f"internal:'Transformer'!A{b3_hdr+1}:Y{b3_end+1}", link_fmt, "Multiphase2wXF")
    ws.write_url(4, 0, f"internal:'Transformer'!A{b4_hdr+1}:AA{b4_end+1}", link_fmt, "Multiphase2wXFMutual")

    # Important notes (rows 8–10)
    ws.merge_range(7, 0, 7, 7, "Important notes:", notes_hdr)
    ws.merge_range(8, 0, 8, 7, "Default order of blocks and columns after row 11 must not change", notes_txt)
    ws.merge_range(9, 0, 9, 7, "One empty row between End of each block and the next block is mandatory; otherwise, empty rows are NOT allowed", notes_txt)

    go_top = "internal:'Transformer'!A1"

    # ---- Block 1: Positive-Sequence 2W (empty) ----
    ws.merge_range(b1_title, 0, b1_title, 2, "Positive-Sequence 2W-Transformer", bold)
    ws.write_url(b1_title, 3, go_top, link_fmt, "Go to Type List")
    ws.write_row(b1_hdr, 0, ["ID","Status","From bus","To bus","R (pu)","Xl (pu)","Gmag (pu)","Bmag (pu)","Ratio W1 (pu)","Ratio W2 (pu)","Phase Shift (deg)"], th)
    ws.merge_range(b1_end, 0, b1_end, 2, "End of Positive-Sequence 2W-Transformer")

    # ---- Block 2: Positive-Sequence 3W (empty) ----
    ws.merge_range(b2_title, 0, b2_title, 2, "Positive-Sequence 3W-Transformer", bold)
    ws.write_url(b2_title, 3, go_top, link_fmt, "Go to Type List")
    ws.write_row(b2_hdr, 0, ["ID","Status","Bus1","Bus2","Bus3","R_12 (pu)","Xl_12 (pu)","R_23 (pu)","Xl_23 (pu)","R_31 (pu)","Xl_31 (pu)","Gmag (pu)","Bmag (pu)","Ratio W1 (pu)","Ratio W2 (pu)","Ratio W3 (pu)","Phase Shift W1 (deg)","Phase Shift W2 (deg)","Phase Shift W3 (deg)"], th)
    ws.merge_range(b2_end, 0, b2_end, 2, "End of Positive-Sequence 3W-Transformer")

    # ---- Block 3: Multiphase 2W (parsed) ----
    ws.merge_range(b3_title, 0, b3_title, 2, "Multiphase 2W-Transformer", bold)
    ws.write_url(b3_title, 3, go_top, link_fmt, "Go to Type List")
    ws.write_row(
        b3_hdr, 0,
        ["ID","Status","Number of phases",
         "Bus1","Bus2","Bus3","V (kV)","S_base (kVA)","Conn. type",
         "Bus1","Bus2","Bus3","V (kV)","S_base (kVA)","Conn. type",
         "Tap 1","Tap 2","Tap 3","Lowest Tap","Highest Tap","Min Range (%)","Max Range (%)",
         "X (pu)","RW1 (pu)","RW2 (pu)"],
        th
    )

    rr = b3_first
    for row in rows_mp2w:
        ws.write(rr, 0, row[0])
        ws.write_number(rr, 1, row[1], int0)
        ws.write_number(rr, 2, row[2], int0)
        ws.write(rr, 3, row[3]); ws.write(rr, 4, row[4]); ws.write(rr, 5, row[5])
        ws.write_number(rr, 6, row[6], num2); ws.write_number(rr, 7, row[7], num2); ws.write(rr, 8, row[8])
        ws.write(rr, 9, row[9]); ws.write(rr,10, row[10]); ws.write(rr,11, row[11])
        ws.write_number(rr,12, row[12], num2); ws.write_number(rr,13, row[13], num2); ws.write(rr,14, row[14])
        ws.write_number(rr,15, row[15], int0); ws.write_number(rr,16, row[16], int0); ws.write_number(rr,17, row[17], int0)
        ws.write_number(rr,18, row[18], int0); ws.write_number(rr,19, row[19], int0)
        ws.write_number(rr,20, row[20], int0); ws.write_number(rr,21, row[21], int0)
        ws.write_number(rr,22, row[22], num5)
        ws.write_number(rr,23, row[23], num8)
        ws.write_number(rr,24, row[24], num8)
        rr += 1

    ws.merge_range(b3_end, 0, b3_end, 3, "End of Multiphase 2W-Transformer")

    # ---- Block 4: Multiphase 2W with Mutual (empty) ----
    ws.merge_range(b4_title, 0, b4_title, 2, "Multiphase 2W-Transformer with Mutual Impedance", bold)
    ws.write_url(b4_title, 3, go_top, link_fmt, "Go to Type List")
    ws.write_row(
        b4_hdr, 0,
        ["ID","Status","Number of phases",
         "Bus1","Bus2","Bus3","V (kV)","S_base (kVA)","Conn. type",
         "Bus1","Bus2","Bus3","V (kV)","S_base (kVA)","Conn. type",
         "Tap 1","Tap 2","Tap 3","Lowest Tap","Highest Tap","Min Range (%)","Max Range (%)",
         "Z0 leakage (pu)","Z1 leakage (pu)","X0/R0","X1/R1","No Load Loss (kW)"],
        th
    )
    ws.merge_range(b4_end, 0, b4_end, 3, "End of Multiphase 2W-Transformer with Mutual Impedance")
