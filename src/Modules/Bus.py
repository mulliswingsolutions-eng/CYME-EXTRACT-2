"""
Bus.py
------
Bus extraction utilities for CYME text export (XML content).

Public API:
    extract_bus_data(path: str | Path) -> list[dict]

Columns returned (where available):
    - Bus
    - Base Voltage (V)
    - Initial Vmag
    - Unit
    - Angle
    - Type
"""

import xml.etree.ElementTree as ET


def extract_bus_data(filepath):
    tree = ET.parse(filepath)
    root = tree.getroot()

    # --- Source info (this is the slack bus) ---
    source_node = (root.findtext('.//Sources/Source/SourceNodeID') or '').strip()

    eq = root.find('.//Sources/Source/EquivalentSourceModels/EquivalentSourceModel/EquivalentSource')
    if eq is None:
        raise ValueError("Could not find EquivalentSource block")

    # per-phase LN kV and angles from the source (OperatingVoltage is LN in kV in this file)
    phase_v_kv = {
        'A': float(eq.findtext('OperatingVoltage1', '0')),
        'B': float(eq.findtext('OperatingVoltage2', '0')),
        'C': float(eq.findtext('OperatingVoltage3', '0')),
    }
    phase_ang = {
        'A': float(eq.findtext('OperatingAngle1', '0')),
        'B': float(eq.findtext('OperatingAngle2', '0')),
        'C': float(eq.findtext('OperatingAngle3', '0')),
    }

    # --- Build node -> phases map by scanning all sections ---
    node_phases = {}  # {node_id: set('A','B','C')}

    def add_phases(node_id, phasestr):
        node_id = (node_id or '').strip()
        if not node_id:
            return
        up = node_id.upper()
        if up.startswith('LOAD') or 'CAP' in up:
            return
        s = node_phases.setdefault(node_id, set())
        for ch in phasestr:
            if ch in ('A', 'B', 'C'):
                s.add(ch)

    for sec in root.findall('.//Sections/Section'):
        phasestr = (sec.findtext('Phase') or '').strip()
        if not phasestr:
            continue
        add_phases(sec.findtext('FromNodeID'), phasestr)
        add_phases(sec.findtext('ToNodeID'), phasestr)

    # Ensure the source node has all three phases if its EquivalentSource is three-phase
    if source_node:
        node_phases.setdefault(source_node, set()).update(['A', 'B', 'C'])

    # --- Emit rows per node-phase ---
    entries = []
    for node_id, phases in sorted(node_phases.items()):
        for ph in sorted(phases):
            v_ln_volts = phase_v_kv[ph] * 1000.0
            entries.append({
                "Bus": f"{node_id}_{ph.lower()}",
                "Base Voltage (V)": v_ln_volts,
                "Initial Vmag": v_ln_volts,  # initial = source LN mag here
                "Unit": "V",
                "Angle": phase_ang[ph],
                "Type": "Slack" if node_id == source_node else "PQ",
            })

    return entries
