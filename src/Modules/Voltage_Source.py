"""
Voltage_Source.py
-----------------
Extract voltage source information from a CYME text export.

Public API:
    extract_voltage_source_data(path: str | Path) -> pandas.DataFrame

Columns:
    - SourceNodeID
    - DeviceNumber
    - SourceID
    - DesiredVoltage_kVLL
    - EquivalentConfig
    - KVLL
    - OperatingVoltage1_kVLN
    - OperatingAngle1_deg
    - PosSeqR
    - PosSeqX
    - NegSeqR
    - NegSeqX
    - ZeroSeqR
    - ZeroSeqX
    - NominalCapacity1_MVA
    - NominalCapacity2_MVA
"""

from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET
from typing import Optional, List, Dict, Any

import pandas as pd


def _first_text(elem: Optional[ET.Element], default: str = "") -> str:
    if elem is None:
        return default
    return (elem.text or "").strip()


def _float_or_none(s: str) -> Optional[float]:
    try:
        return float(s)
    except Exception:
        return None


def _get_network(root: ET.Element) -> Optional[ET.Element]:
    nets = root.find("Networks")
    if nets is None:
        return None
    return nets.find("Network")


def extract_voltage_source_data(path: str | Path) -> pd.DataFrame:
    """
    Parse source information from a CYME text export.

    Parameters
    ----------
    path : str | Path

    Returns
    -------
    pandas.DataFrame
    """
    p = Path(path)
    tree = ET.parse(p)
    root = tree.getroot()

    network = _get_network(root)
    if network is None:
        return pd.DataFrame(
            columns=[
                "SourceNodeID",
                "DeviceNumber",
                "SourceID",
                "DesiredVoltage_kVLL",
                "EquivalentConfig",
                "KVLL",
                "OperatingVoltage1_kVLN",
                "OperatingAngle1_deg",
                "PosSeqR",
                "PosSeqX",
                "NegSeqR",
                "NegSeqX",
                "ZeroSeqR",
                "ZeroSeqX",
                "NominalCapacity1_MVA",
                "NominalCapacity2_MVA",
            ]
        )

    topos = network.find("Topos")
    topo = topos.find("Topo") if topos is not None else None
    sources_parent = topo.find("Sources") if topo is not None else None

    rows: List[Dict[str, Any]] = []

    if sources_parent is not None:
        for src in sources_parent.findall("Source"):
            src_node_id = _first_text(src.find("SourceNodeID"))
            settings = src.find("SourceSettings")
            device_number = _first_text(settings.find("DeviceNumber")) if settings is not None else ""
            source_id = _first_text(settings.find("SourceID")) if settings is not None else ""
            desired_v = _float_or_none(_first_text(settings.find("DesiredVoltage"))) if settings is not None else None

            eq_config = _first_text(src.find("EquivalentSourceConfiguration"))
            eq_models = src.find("EquivalentSourceModels")
            kvll = op_v1 = op_ang1 = None
            r1 = x1 = r2 = x2 = r0 = x0 = None
            nom1 = nom2 = None

            if eq_models is not None:
                eq_model = eq_models.find("EquivalentSourceModel")
                if eq_model is not None:
                    eq_src = eq_model.find("EquivalentSource")
                    if eq_src is not None:
                        kvll = _float_or_none(_first_text(eq_src.find("KVLL")))
                        op_v1 = _float_or_none(_first_text(eq_src.find("OperatingVoltage1")))
                        op_ang1 = _float_or_none(_first_text(eq_src.find("OperatingAngle1")))
                        r1 = _float_or_none(_first_text(eq_src.find("PositiveSequenceResistance")))
                        x1 = _float_or_none(_first_text(eq_src.find("PositiveSequenceReactance")))
                        r2 = _float_or_none(_first_text(eq_src.find("NegativeSequenceResistance")))
                        x2 = _float_or_none(_first_text(eq_src.find("NegativeSequenceReactance")))
                        r0 = _float_or_none(_first_text(eq_src.find("ZeroSequenceResistance")))
                        x0 = _float_or_none(_first_text(eq_src.find("ZeroSequenceReactance")))
                        nom1 = _float_or_none(_first_text(eq_src.find("NominalCapacity1MVA")))
                        nom2 = _float_or_none(_first_text(eq_src.find("NominalCapacity2MVA")))

            rows.append(
                {
                    "SourceNodeID": src_node_id,
                    "DeviceNumber": device_number,
                    "SourceID": source_id,
                    "DesiredVoltage_kVLL": desired_v,
                    "EquivalentConfig": eq_config,
                    "KVLL": kvll,
                    "OperatingVoltage1_kVLN": op_v1,
                    "OperatingAngle1_deg": op_ang1,
                    "PosSeqR": r1,
                    "PosSeqX": x1,
                    "NegSeqR": r2,
                    "NegSeqX": x2,
                    "ZeroSeqR": r0,
                    "ZeroSeqX": x0,
                    "NominalCapacity1_MVA": nom1,
                    "NominalCapacity2_MVA": nom2,
                }
            )

    df = pd.DataFrame(
        rows,
        columns=[
            "SourceNodeID",
            "DeviceNumber",
            "SourceID",
            "DesiredVoltage_kVLL",
            "EquivalentConfig",
            "KVLL",
            "OperatingVoltage1_kVLN",
            "OperatingAngle1_deg",
            "PosSeqR",
            "PosSeqX",
            "NegSeqR",
            "NegSeqX",
            "ZeroSeqR",
            "ZeroSeqX",
            "NominalCapacity1_MVA",
            "NominalCapacity2_MVA",
        ],
    )

    return df
