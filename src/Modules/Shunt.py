# Modules/Shunt.py
from __future__ import annotations
from pathlib import Path
import xml.etree.ElementTree as ET
from typing import List, Dict, Any
from Modules.General import safe_name
from Modules.IslandFilter import should_comment_bus

PHASE_SUFFIX = {"A": "_a", "B": "_b", "C": "_c"}
PHASES = ("A", "B", "C")

def _cap_id(from_node: str, device_number: str) -> str:
    """
    Goal style prefers 'cap<digits>' (e.g., cap611, cap675).
    We extract digits from the sanitized FromNodeID or DeviceNumber.
    If none, fall back to a sanitized identifier.
    """
    fn = safe_name(from_node)
    dn = safe_name(device_number)
    s = "".join(ch for ch in (fn or dn or "") if ch.isdigit())
    return f"cap{s}" if s else (dn or fn or "cap")


def _parse_shunts(txt_path: Path):
    """
    Returns (single_rows, two_rows, three_rows)

    single_rows: [ID, Status, kVLN, Bus1, P1kW, Q1kVAr]
    two_rows:    [ID, Status1, Status2, kVLN, Bus1, Bus2, P1, Q1, P2, Q2]
    three_rows:  [ID, Status1, Status2, Status3, kVLN, Bus1, Bus2, Bus3, P1, Q1, P2, Q2, P3, Q3]

    NEW: If the shunt's bus base is not active on the Bus sheet,
         the row is commented by prefixing '//' to the ID.
    """
    root = ET.fromstring(txt_path.read_text(encoding="utf-8", errors="ignore"))

    single_rows: List[List[Any]] = []
    two_rows: List[List[Any]] = []
    three_rows: List[List[Any]] = []

    for sec in root.findall(".//Section"):
        dev = sec.find(".//Devices/ShuntCapacitor")
        if dev is None:
            continue

        from_bus_raw = (sec.findtext("FromNodeID") or "").strip()
        from_bus = safe_name(from_bus_raw)

        phase = (sec.findtext("Phase") or "ABC").strip().upper()

        devnum_raw = (dev.findtext("DeviceNumber") or "").strip()
        devnum = safe_name(devnum_raw)

        status = 1 if (dev.findtext("ConnectionStatus") or "Connected").strip().lower() == "connected" else 0
        kvln = dev.findtext("KVLN")
        kvln = float(kvln) if kvln not in (None, "") else None

        # Fixed losses (kW) and kVAr per phase
        def _f(txt: str | None) -> float:
            try:
                return float(txt) if txt not in (None, "") else 0.0
            except Exception:
                return 0.0

        kW = {
            "A": _f(dev.findtext("FixedLossesA")),
            "B": _f(dev.findtext("FixedLossesB")),
            "C": _f(dev.findtext("FixedLossesC")),
        }

        kVAr = {
            "A": _f(dev.findtext("FixedKVARA")),
            "B": _f(dev.findtext("FixedKVARB")),
            "C": _f(dev.findtext("FixedKVARC")),
        }

        # Convention: capacitors inject negative Q in the sheet
        for p in PHASES:
            kVAr[p] = -kVAr[p]

        # Build ID and maybe comment it out per island policy
        cid = _cap_id(from_bus, devnum)
        cid_out = cid if not should_comment_bus(from_bus) else f"//{cid}"

        if phase in PHASES:
            p = phase
            single_rows.append([
                cid_out, status, kvln,
                f"{from_bus}{PHASE_SUFFIX[p]}",
                kW[p], kVAr[p]
            ])

        elif phase in ("AB", "BC", "AC"):
            p1, p2 = phase[0], phase[1]
            two_rows.append([
                cid_out, status, status, kvln,
                f"{from_bus}{PHASE_SUFFIX[p1]}", f"{from_bus}{PHASE_SUFFIX[p2]}",
                kW[p1], kVAr[p1], kW[p2], kVAr[p2]
            ])

        else:  # treat everything else as three-phase (ABC)
            three_rows.append([
                cid_out, status, status, status, kvln,
                f"{from_bus}_a", f"{from_bus}_b", f"{from_bus}_c",
                kW["A"], kVAr["A"], kW["B"], kVAr["B"], kW["C"], kVAr["C"]
            ])

    # stable sort by numeric part of ID so it matches expectation
    def _key(row):  # row[0] is ID like 'cap675' (possibly prefixed with '//')
        id_clean = row[0][2:] if str(row[0]).startswith("//") else str(row[0])
        digits = "".join(ch for ch in (id_clean or "") if ch.isdigit())
        return int(digits) if digits else 0

    single_rows.sort(key=_key)
    two_rows.sort(key=_key)
    three_rows.sort(key=_key)

    return single_rows, two_rows, three_rows


def write_shunt_sheet(xw, input_path: Path) -> None:
    """
    Build the 'Shunt' sheet with the exact layout you specified.
    """
    wb = xw.book
    ws = wb.add_worksheet("Shunt")
    xw.sheets["Shunt"] = ws

    # Formats
    bold = wb.add_format({"bold": True})
    link_fmt = wb.add_format({"font_color": "blue", "underline": 1})
    notes_hdr = wb.add_format({"bold": True, "font_color": "yellow", "bg_color": "#595959", "align": "left"})
    notes_txt = wb.add_format({"font_color": "yellow", "bg_color": "#595959"})
    th = wb.add_format({"bold": True, "bottom": 1})
    num4 = wb.add_format({"num_format": "0.0000"})
    int0 = wb.add_format({"num_format": "0"})

    # Column widths
    widths = [24, 10, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14]
    for c, w in enumerate(widths):
        ws.set_column(c, c, w)

    # Parse data (already sanitized + protected)
    single_rows, two_rows, three_rows = _parse_shunts(input_path)

    # Anchor rows (blocks start at row 11 -> index 10)
    r = 10
    # Positive-Sequence (empty)
    b1_t, b1_h, b1_e = r, r + 1, r + 2; r = b1_e + 2
    # Single-phase
    b2_t, b2_h, b2_first = r, r + 1, r + 2
    b2_e = b2_first + len(single_rows); r = b2_e + 2
    # Two-phase
    b3_t, b3_h, b3_first = r, r + 1, r + 2
    b3_e = b3_first + len(two_rows); r = b3_e + 2
    # Three-phase
    b4_t, b4_h, b4_first = r, r + 1, r + 2
    b4_e = b4_first + len(three_rows)

    # Type row + links
    ws.write(0, 0, "Type", bold)
    ws.write_url(1, 0, f"internal:'Shunt'!A{b1_h+1}:E{b1_e+1}", link_fmt, "PositiveSeqShunt")
    ws.write_url(2, 0, f"internal:'Shunt'!A{b2_h+1}:F{b2_e+1}", link_fmt, "SinglePhaseShunt")
    ws.write_url(3, 0, f"internal:'Shunt'!A{b3_h+1}:J{b3_e+1}", link_fmt, "TwoPhaseShunt")
    ws.write_url(4, 0, f"internal:'Shunt'!A{b4_h+1}:N{b4_e+1}", link_fmt, "ThreePhaseShunt")

    # Notes
    ws.merge_range(7, 0, 7, 7, "Important notes:", notes_hdr)
    ws.merge_range(8, 0, 8, 7, "Default order of blocks and columns after row 11 must not change", notes_txt)
    ws.merge_range(9, 0, 9, 7, "One empty row between End of each block and the next block is mandatory; otherwise, empty rows are NOT allowed", notes_txt)

    go_top = "internal:'Shunt'!A1"

    # ---- Block 1: Positive-Sequence (empty) ----
    ws.merge_range(b1_t, 0, b1_t, 1, "Positive-Sequence Shunt", bold)
    ws.write_url(b1_t, 2, go_top, link_fmt, "Go to Type List")
    ws.write_row(b1_h, 0, ["ID", "Status", "Bus", "P (MW)", "Q (MVAr)"], th)
    ws.merge_range(b1_e, 0, b1_e, 2, "End of Positive-Sequence Shunt")

    # ---- Block 2: Single-Phase (parsed) ----
    ws.merge_range(b2_t, 0, b2_t, 1, "Single-Phase Shunt", bold)
    ws.write_url(b2_t, 2, go_top, link_fmt, "Go to Type List")
    ws.write_row(b2_h, 0, ["ID", "Status", "kV (ph-gr RMS)", "Bus1", "P1 (kW)", "Q1 (kVAr)"], th)
    rr = b2_first
    for row in single_rows:
        ws.write(rr, 0, row[0])
        ws.write_number(rr, 1, row[1], int0)
        if row[2] is None:
            ws.write(rr, 2, "")
        else:
            ws.write_number(rr, 2, row[2], num4)
        ws.write(rr, 3, row[3])
        ws.write_number(rr, 4, row[4], int0)
        ws.write_number(rr, 5, row[5], int0)
        rr += 1
    ws.merge_range(b2_e, 0, b2_e, 2, "End of Single-Phase Shunt")

    # ---- Block 3: Two-Phase (parsed) ----
    ws.merge_range(b3_t, 0, b3_t, 1, "Two-Phase Shunt", bold)
    ws.write_url(b3_t, 2, go_top, link_fmt, "Go to Type List")
    ws.write_row(b3_h, 0, ["ID", "Status1", "Status2", "kV (ph-gr RMS)", "Bus1", "Bus2", "P1 (kW)", "Q1 (kVAr)", "P2 (kW)", "Q2 (kVAr)"], th)
    rr = b3_first
    for row in two_rows:
        ws.write(rr, 0, row[0])
        ws.write_number(rr, 1, row[1], int0)
        ws.write_number(rr, 2, row[2], int0)
        if row[3] is None:
            ws.write(rr, 3, "")
        else:
            ws.write_number(rr, 3, row[3], num4)
        ws.write(rr, 4, row[4]); ws.write(rr, 5, row[5])
        ws.write_number(rr, 6, row[6], int0); ws.write_number(rr, 7, row[7], int0)
        ws.write_number(rr, 8, row[8], int0); ws.write_number(rr, 9, row[9], int0)
        rr += 1
    ws.merge_range(b3_e, 0, b3_e, 2, "End of Two-Phase Shunt")

    # ---- Block 4: Three-Phase (parsed) ----
    ws.merge_range(b4_t, 0, b4_t, 1, "Three-Phase Shunt", bold)
    ws.write_url(b4_t, 2, go_top, link_fmt, "Go to Type List")
    ws.write_row(b4_h, 0, ["ID", "Status1", "Status2", "Status3", "kV (ph-gr RMS)", "Bus1", "Bus2", "Bus3", "P1 (kW)", "Q1 (kVAr)", "P2 (kW)", "Q2 (kVAr)", "P3 (kW)", "Q3 (kVAr)"], th)
    rr = b4_first
    for row in three_rows:
        ws.write(rr, 0, row[0])
        ws.write_number(rr, 1, row[1], int0); ws.write_number(rr, 2, row[2], int0); ws.write_number(rr, 3, row[3], int0)
        if row[4] is None:
            ws.write(rr, 4, "")
        else:
            ws.write_number(rr, 4, row[4], num4)
        ws.write(rr, 5, row[5]); ws.write(rr, 6, row[6]); ws.write(rr, 7, row[7])
        ws.write_number(rr, 8,  row[8],  int0); ws.write_number(rr, 9,  row[9],  int0)
        ws.write_number(rr, 10, row[10], int0); ws.write_number(rr, 11, row[11], int0)
        ws.write_number(rr, 12, row[12], int0); ws.write_number(rr, 13, row[13], int0)
        rr += 1
    ws.merge_range(b4_e, 0, b4_e, 2, "End of Three-Phase Shunt")
