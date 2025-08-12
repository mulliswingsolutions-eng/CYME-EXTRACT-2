# Modules/Bus.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple
import xml.etree.ElementTree as ET

PHASES = ("A", "B", "C")


def _parse_bus_rows(input_path: Path) -> List[Dict]:
    """
    Build rows for the Bus sheet.

    Columns:
      Bus | Base Voltage (V) | Initial Vmag | Unit | Angle | Type
    """
    # Parse the XML text (CYME export)
    root = ET.fromstring(Path(input_path).read_text(encoding="utf-8", errors="ignore"))

    # --- Slack/source info ---
    source_node = (root.findtext(".//Sources/Source/SourceNodeID") or "").strip()

    eq = root.find(".//Sources/Source/EquivalentSourceModels/EquivalentSourceModel/EquivalentSource")
    if eq is None:
        # keep robust: no rows if no source found
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

    # --- Build node -> phases map by scanning all sections ---
    node_phases: Dict[str, set] = {}

    def add_phases(node_id: str | None, phase_str: str | None) -> None:
        node_id = (node_id or "").strip()
        if not node_id:
            return
        up = node_id.upper()
        # Exclude device pseudo-nodes (LOAD*, *CAP*)
        if up.startswith("LOAD") or "CAP" in up:
            return
        s = node_phases.setdefault(node_id, set())
        for ch in (phase_str or ""):
            if ch in PHASES:
                s.add(ch)

    for sec in root.findall(".//Sections/Section"):
        phasestr = (sec.findtext("Phase") or "").strip()
        if not phasestr:
            continue
        add_phases(sec.findtext("FromNodeID"), phasestr)
        add_phases(sec.findtext("ToNodeID"), phasestr)

    # Ensure the slack/source node has all three phases (typical 3φ source)
    if source_node:
        node_phases.setdefault(source_node, set()).update(PHASES)

    # --- Emit rows per node-phase ---
    def phase_key(p: str) -> int:
        return PHASES.index(p)

    rows: List[Dict] = []
    for node_id in sorted(node_phases.keys()):
        for ph in sorted(node_phases[node_id], key=phase_key):
            v_ln_volts = phase_v_kv[ph] * 1000.0
            rows.append(
                {
                    "Bus": f"{node_id}_{ph.lower()}",
                    "Base Voltage (V)": v_ln_volts,
                    "Initial Vmag": v_ln_volts,  # use source LN magnitude as initial
                    "Unit": "V",
                    "Angle": phase_ang[ph],
                    "Type": "Slack" if node_id == source_node else "PQ",
                }
            )

    return rows


# --- Backward compatibility (original API) ---
def extract_bus_data(filepath: str | Path) -> List[Dict]:
    return _parse_bus_rows(Path(filepath))


# --- Unified writer API (matches other modules) ---
def write_bus_sheet(xw, input_path: Path) -> None:
    """
    Create the 'Bus' sheet using xlsxwriter (same pattern as other pages).
    """
    rows = _parse_bus_rows(Path(input_path))

    wb = xw.book
    ws = wb.add_worksheet("Bus")
    # expose for consistency if you keep a sheets dict on the writer
    try:
        xw.sheets["Bus"] = ws
    except Exception:
        pass

    # Formats
    hdr = wb.add_format({"bold": True})
    num0 = wb.add_format({"num_format": "0"})
    num2 = wb.add_format({"num_format": "0.00"})

    # Column widths
    ws.set_column(0, 0, 18)  # Bus
    ws.set_column(1, 1, 18)  # Base Voltage (V)
    ws.set_column(2, 2, 16)  # Initial Vmag
    ws.set_column(3, 3, 8)   # Unit
    ws.set_column(4, 4, 10)  # Angle
    ws.set_column(5, 5, 10)  # Type

    # Header
    headers = ["Bus", "Base Voltage (V)", "Initial Vmag", "Unit", "Angle", "Type"]
    ws.write_row(0, 0, headers, hdr)

    # Data
    r = 1
    for row in rows:
        ws.write(r, 0, row["Bus"])
        ws.write_number(r, 1, float(row["Base Voltage (V)"]), num0)
        ws.write_number(r, 2, float(row["Initial Vmag"]), num0)
        ws.write(r, 3, row["Unit"])
        ws.write_number(r, 4, float(row["Angle"]), num2)
        ws.write(r, 5, row["Type"])
        r += 1
