# src/main.py
from pathlib import Path
import xml.etree.ElementTree as ET
import pandas as pd

INPUT_SXST = r"Examples\Example-13bus-modified.txt"
OUTPUT_XLSX = r"Outputs\CYME_Extract_13Bus.xlsx"

def text(el, default=""):
    return el.text.strip() if el is not None and el.text is not None else default

def get(root, path):
    el = root.find(path)
    return text(el)

def all_elems(root, path):
    return root.findall(path)

def parse_nodes(root):
    rows = []
    for n in all_elems(root, ".//Nodes/Node"):
        rows.append({
            "NodeID": get(n, "./NodeID"),
            "UserDefinedBaseVoltage": get(n, "./UserDefinedBaseVoltage"),
        })
    return pd.DataFrame(rows)

def parse_sections(root):
    rows = []
    for s in all_elems(root, ".//Sections/Section"):
        rows.append({
            "SectionID": get(s, "./SectionID"),
            "FromNodeID": get(s, "./FromNodeID"),
            "ToNodeID": get(s, "./ToNodeID"),
            "Phase": get(s, "./Phase"),
        })
    return pd.DataFrame(rows)

def parse_lines(root):
    rows = []
    for s in all_elems(root, ".//Sections/Section"):
        # OverheadByPhase
        for dev in all_elems(s, "./Devices/OverheadByPhase"):
            rows.append({
                "SectionID": get(s, "./SectionID"),
                "LineType": "OverheadByPhase",
                "Length": get(dev, "./Length"),
                "ConductorPosition": get(dev, "./ConductorPosition"),
                "PhaseA": get(dev, "./PhaseConductorIDA"),
                "PhaseB": get(dev, "./PhaseConductorIDB"),
                "PhaseC": get(dev, "./PhaseConductorIDC"),
                "Neutral1": get(dev, "./NeutralConductorID1"),
                "Neutral2": get(dev, "./NeutralConductorID2"),
                "SpacingID": get(dev, "./ConductorSpacingID"),
                "EarthResistivity": get(dev, "./EarthResistivity"),
            })
        # OverheadLineUnbalanced
        for dev in all_elems(s, "./Devices/OverheadLineUnbalanced"):
            rows.append({
                "SectionID": get(s, "./SectionID"),
                "LineType": "OverheadLineUnbalanced",
                "Length": get(dev, "./Length"),
                "LineID": get(dev, "./LineID"),
                "ConductorPosition": "",
                "PhaseA": "",
                "PhaseB": "",
                "PhaseC": "",
                "Neutral1": "",
                "Neutral2": "",
                "SpacingID": "",
                "EarthResistivity": "",
            })
    return pd.DataFrame(rows)

def _per_phase_load_values(dev):
    # Collect KW/KVAR per phase if present
    ph = {"KW_A": "", "KVAR_A": "", "KW_B": "", "KVAR_B": "", "KW_C": "", "KVAR_C": ""}
    for lv in all_elems(dev, ".//CustomerLoadValues/CustomerLoadValue"):
        phase = get(lv, "./Phase").upper()
        KW = get(lv, "./LoadValue/KW")
        KVAR = get(lv, "./LoadValue/KVAR")
        if phase == "A":
            ph["KW_A"], ph["KVAR_A"] = KW, KVAR
        elif phase == "B":
            ph["KW_B"], ph["KVAR_B"] = KW, KVAR
        elif phase == "C":
            ph["KW_C"], ph["KVAR_C"] = KW, KVAR
    return ph

def parse_spot_loads(root):
    rows = []
    for s in all_elems(root, ".//Sections/Section"):
        for dev in all_elems(s, "./Devices/SpotLoad"):
            ph = _per_phase_load_values(dev)
            rows.append({
                "SectionID": get(s, "./SectionID"),
                "FromNodeID": get(s, "./FromNodeID"),
                "ToNodeID": get(s, "./ToNodeID"),
                "DeviceNumber": get(dev, "./DeviceNumber"),
                "ConnConfig": get(dev, "./ConnectionConfiguration"),
                **ph
            })
    return pd.DataFrame(rows)

def parse_dist_loads(root):
    rows = []
    for s in all_elems(root, ".//Sections/Section"):
        for dev in all_elems(s, "./Devices/DistributedLoad"):
            ph = _per_phase_load_values(dev)
            rows.append({
                "SectionID": get(s, "./SectionID"),
                "FromNodeID": get(s, "./FromNodeID"),
                "ToNodeID": get(s, "./ToNodeID"),
                "DeviceNumber": get(dev, "./DeviceNumber"),
                "ConnConfig": get(dev, "./ConnectionConfiguration"),
                **ph
            })
    return pd.DataFrame(rows)

def parse_caps(root):
    rows = []
    for s in all_elems(root, ".//Sections/Section"):
        for dev in all_elems(s, "./Devices/ShuntCapacitor"):
            rows.append({
                "SectionID": get(s, "./SectionID"),
                "FromNodeID": get(s, "./FromNodeID"),
                "ToNodeID": get(s, "./ToNodeID"),
                "DeviceNumber": get(dev, "./DeviceNumber"),
                "ConnConfig": get(dev, "./ConnectionConfiguration"),
                "KVAR_A": get(dev, "./FixedKVARA"),
                "KVAR_B": get(dev, "./FixedKVARB"),
                "KVAR_C": get(dev, "./FixedKVARC"),
                "KVLN": get(dev, "./KVLN"),
            })
    return pd.DataFrame(rows)

def parse_regs(root):
    rows = []
    for s in all_elems(root, ".//Sections/Section"):
        for dev in all_elems(s, "./Devices/Regulator"):
            rows.append({
                "SectionID": get(s, "./SectionID"),
                "FromNodeID": get(s, "./FromNodeID"),
                "ToNodeID": get(s, "./ToNodeID"),
                "DeviceNumber": get(dev, "./DeviceNumber"),
                "ConnConfig": get(dev, "./ConnectionConfiguration"),
                "TapA": get(dev, "./TapPositionA"),
                "TapB": get(dev, "./TapPositionB"),
                "TapC": get(dev, "./TapPositionC"),
                "BandWidth": get(dev, "./BandWidth"),
                "BoostPercent": get(dev, "./BoostPercent"),
                "BuckPercent": get(dev, "./BuckPercent"),
            })
    return pd.DataFrame(rows)

def parse_transformers(root):
    rows = []
    for s in all_elems(root, ".//Sections/Section"):
        for dev in all_elems(s, "./Devices/Transformer"):
            rows.append({
                "SectionID": get(s, "./SectionID"),
                "FromNodeID": get(s, "./FromNodeID"),
                "ToNodeID": get(s, "./ToNodeID"),
                "DeviceNumber": get(dev, "./DeviceNumber"),
                "DeviceID": get(dev, "./DeviceID"),
                "Connection": get(dev, "./TransformerConnection"),
                "PhaseShift": get(dev, "./PhaseShift"),
                "PrimaryTapPct": get(dev, "./PrimaryTapSettingPercent"),
                "SecondaryTapPct": get(dev, "./SecondaryTapSettingPercent"),
            })
    return pd.DataFrame(rows)

def parse_switches(root):
    rows = []
    for s in all_elems(root, ".//Sections/Section"):
        for dev in all_elems(s, "./Devices/Switch"):
            rows.append({
                "SectionID": get(s, "./SectionID"),
                "FromNodeID": get(s, "./FromNodeID"),
                "ToNodeID": get(s, "./ToNodeID"),
                "DeviceNumber": get(dev, "./DeviceNumber"),
                "DeviceID": get(dev, "./DeviceID"),
                "ClosedPhase": get(dev, "./ClosedPhase"),
                "NormalStatus": get(dev, "./NormalStatus"),
                "RemoteControlled": get(dev, "./RemoteControlled"),
            })
    return pd.DataFrame(rows)

def parse_sources(root):
    rows = []
    for src in all_elems(root, ".//Sources/Source"):
        sset = src.find("./SourceSettings")
        eq = src.find(".//EquivalentSourceModels/EquivalentSourceModel/EquivalentSource")
        rows.append({
            "SourceNodeID": get(src, "./SourceNodeID"),
            "SourceID": get(sset, "./SourceID") if sset is not None else "",
            "DesiredKVLL": get(sset, "./DesiredVoltage"),
            "KVLL": get(eq, "./KVLL") if eq is not None else "",
            "PosSeqR": get(eq, "./PositiveSequenceResistance") if eq is not None else "",
            "PosSeqX": get(eq, "./PositiveSequenceReactance") if eq is not None else "",
            "ZeroSeqR": get(eq, "./ZeroSequenceResistance") if eq is not None else "",
            "ZeroSeqX": get(eq, "./ZeroSequenceReactance") if eq is not None else "",
        })
    return pd.DataFrame(rows)

def main():
    in_path = Path(INPUT_SXST)
    out_path = Path(OUTPUT_XLSX)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    tree = ET.parse(in_path)
    root = tree.getroot()

    dfs = {
        "Nodes": parse_nodes(root),
        "Sections": parse_sections(root),
        "Lines": parse_lines(root),
        "Loads": parse_spot_loads(root),
        "DistLoads": parse_dist_loads(root),
        "Caps": parse_caps(root),
        "Regulators": parse_regs(root),
        "Transformers": parse_transformers(root),
        "Switches": parse_switches(root),
        "Sources": parse_sources(root),
    }

    with pd.ExcelWriter(out_path, engine="openpyxl") as xl:
        for name, df in dfs.items():
            # Keep stable column order if possible
            if not df.empty:
                df.to_excel(xl, sheet_name=name, index=False)
            else:
                # write headers-only empty sheet to keep structure
                pd.DataFrame(columns=[]).to_excel(xl, sheet_name=name, index=False)

    print(f"✅ Wrote {out_path}")

if __name__ == "__main__":
    main()
