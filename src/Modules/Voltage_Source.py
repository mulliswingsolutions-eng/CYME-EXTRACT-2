"""
Voltage_Source.py
-----------------
Voltage source extraction utilities for CYME text export (XML content).

Public API:
    extract_voltage_source_data(path: str | Path) -> pandas.DataFrame

The returned DataFrame matches the Excel "Voltage Source" sheet template exactly.

Blocks included:
    - Positive-Sequence Voltage Source (empty for now)
    - Single-Phase Voltage Source (empty for now)
    - Three-Phase Voltage Source with Short-Circuit Level Data (from <Sources>)
    - Three-Phase Voltage Source with Sequential Data (empty for now)

Important notes:
    - Default order of blocks and columns after row 11 must not change.
    - One empty row between end of each block and the next block is mandatory.
    - Other empty rows are not allowed.
"""

from pathlib import Path
import pandas as pd
import xml.etree.ElementTree as ET


def extract_voltage_source_data(path: str | Path) -> pd.DataFrame:
    # Read file and parse XML
    with open(path, "r", encoding="utf-8") as f:
        xml_content = f.read()
    root = ET.fromstring(xml_content)

    # --- Block 1: Positive-Sequence Voltage Source (empty) ---
    positive_seq_header = [
        ["Positive-Sequence Voltage Source", "Go to Type List"],
        ["ID", "Bus", "Voltage (pu)", "Angle (deg)", "Rs (pu)", "Xs (pu)"],
        ["End of Positive-Sequence Voltage Source"],
        [""],  # Empty row after block
    ]

    # --- Block 2: Single-Phase Voltage Source (empty) ---
    single_phase_header = [
        ["Single-Phase Voltage Source", "Go to Type List"],
        ["ID", "Bus1", "Voltage (V)", "Angle (deg)", "Rs (Ohm)", "Xs (Ohm)"],
        ["End of Single-Phase Voltage Source"],
        [""],  # Empty row after block
    ]

    # --- Block 3: Three-Phase Voltage Source with Short-Circuit Level Data ---
    short_circuit_header = [
        ["Three-Phase Voltage Source with Short-Circuit Level Data", "Go to Type List"],
        [
            "ID", "Bus1", "Bus2", "Bus3",
            "kV (ph-ph RMS)", "Angle_a (deg)",
            "SC1ph (MVA)", "SC3ph (MVA)"
        ]
    ]
    short_circuit_rows = []

    sources_elem = root.find(".//Sources")
    if sources_elem is not None:
        for src in sources_elem.findall(".//Source"):
            node_id = src.findtext(".//SourceNodeID")
            src_id = src.findtext(".//SourceID")
            kvll = src.findtext(".//KVLL")
            angle_a = src.findtext(".//OperatingAngle1")

            if node_id and src_id and kvll and angle_a:
                short_circuit_rows.append([
                    src_id,
                    f"{node_id}_a",
                    f"{node_id}_b",
                    f"{node_id}_c",
                    float(kvll),
                    float(angle_a),
                    200000,  # SC1ph hardcoded
                    200000   # SC3ph hardcoded
                ])

    short_circuit_footer = [["End of Three-Phase Voltage Source Short-Circuit Level Data"], [""]]

    # --- Block 4: Three-Phase Voltage Source with Sequential Data (empty) ---
    sequential_header = [
        ["Three-Phase Voltage Source with Sequential Data", "Go to Type List"],
        [
            "ID", "Bus1", "Bus2", "Bus3",
            "kV (ph-ph RMS)", "Angle_a (deg)",
            "R1 (Ohm)", "X1 (Ohm)", "R0 (Ohm)", "X0 (Ohm)"
        ],
        ["End of Three-Phase Voltage Source Sequential Data"]
    ]

    # --- Combine all blocks into one DataFrame ---
    data = (
        positive_seq_header +
        single_phase_header +
        short_circuit_header + short_circuit_rows + short_circuit_footer +
        sequential_header
    )

    # Pad rows so all have same number of columns (Excel formatting)
    max_cols = max(len(row) for row in data)
    data = [row + [""] * (max_cols - len(row)) for row in data]

    return pd.DataFrame(data)
