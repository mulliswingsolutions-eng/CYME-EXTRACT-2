# Modules/Bus.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List
import xml.etree.ElementTree as ET

PHASES = ("A", "B", "C")


def _parse_bus_rows(input_path: Path) -> List[Dict]:
    """
    Build rows for the Bus sheet.

    Columns:
      Bus | Base Voltage (V) | Initial Vmag | Unit | Angle | Type
    """
    root = ET.fromstring(Path(input_path).read_text(encoding="utf-8", errors="ignore"))

    # --- Slack/source info ---
    source_node = (root.findtext(".//Sources/Source/SourceNodeID") or "").strip()

    eq = root.find(".//Sources/Source/EquivalentSourceModels/EquivalentSourceModel/EquivalentSource")
    if eq is None:
        return []

    # LN kV and angles from the source
    phase_v_kv = {
        "A": float(eq.findtext("OperatingVoltage1", "0")),
        "B": float(eq.findtext("OperatingVoltage2", "0")),
        "C": float(eq.findtext("OperatingVoltage3", "0")),
    }
    phase_ang = {
        "A": float(eq.findtext("OperatingAngle1", "0")),
        "B": float(eq.findtext("OperatingAngle2", "0")),
        "C": float(eq.findtext("OperatingAngle3", "0")),
    }

    # --- Build node -> phases map while avoiding device terminals ---
    node_phases: Dict[str, set] = {}

    def add(node_id: str | None, phase_str: str) -> None:
        nid = (node_id or "").strip()
        if not nid:
            return
        s = node_phases.setdefault(nid, set())
        for ch in phase_str:
            if ch in PHASES:
                s.add(ch)

    for sec in root.findall(".//Sections/Section"):
        ph = (sec.findtext("Phase") or "ABC").strip().upper()

        has_line   = sec.find(".//Devices/OverheadLineUnbalanced") is not None or \
                     sec.find(".//Devices/OverheadByPhase") is not None
        has_xfmr   = sec.find(".//Devices/Transformer") is not None
        has_switch = sec.find(".//Devices/Switch") is not None
        has_spot   = sec.find(".//Devices/SpotLoad") is not None
        has_shunt  = sec.find(".//Devices/ShuntCapacitor") is not None or \
                     sec.find(".//Devices/ShuntReactor") is not None

        f = sec.findtext("FromNodeID")
        t = sec.findtext("ToNodeID")

        if has_line or has_xfmr or has_switch:
            # Network branches use both endpoints as buses
            add(f, ph)
            add(t, ph)
        elif has_spot or has_shunt:
            # Device sections: only the FromNodeID is a real bus
            add(f, ph)

    if source_node:
        node_phases.setdefault(source_node, set()).update(PHASES)

    # --- Emit rows per node-phase ---
    def phase_key(p: str) -> int:
        return PHASES.index(p)

    rows: List[Dict] = []
    for node in sorted(node_phases):
        for ph in sorted(node_phases[node], key=phase_key):
            v_ln = phase_v_kv[ph] * 1000.0
            rows.append({
                "Bus": f"{node}_{ph.lower()}",
                "Base Voltage (V)": v_ln,
                "Initial Vmag": v_ln,
                "Unit": "V",
                "Angle": phase_ang[ph],
                "Type": "Slack" if node == source_node else "PQ",
            })
    return rows


# Backward-compat API
def extract_bus_data(filepath: str | Path) -> List[Dict]:
    return _parse_bus_rows(Path(filepath))


# Unified writer API
def write_bus_sheet(xw, input_path: Path) -> None:
    rows = _parse_bus_rows(Path(input_path))

    wb = xw.book
    ws = wb.add_worksheet("Bus")
    xw.sheets["Bus"] = ws

    hdr = wb.add_format({"bold": True})
    num0 = wb.add_format({"num_format": "0"})
    num2 = wb.add_format({"num_format": "0.00"})

    ws.set_column(0, 0, 18)
    ws.set_column(1, 2, 18)
    ws.set_column(3, 3, 8)
    ws.set_column(4, 5, 10)

    ws.write_row(0, 0, ["Bus", "Base Voltage (V)", "Initial Vmag", "Unit", "Angle", "Type"], hdr)
    r = 1
    for row in rows:
        ws.write(r, 0, row["Bus"])
        ws.write_number(r, 1, float(row["Base Voltage (V)"]), num0)
        ws.write_number(r, 2, float(row["Initial Vmag"]), num0)
        ws.write(r, 3, row["Unit"])
        ws.write_number(r, 4, float(row["Angle"]), num2)
        ws.write(r, 5, row["Type"])
        r += 1
