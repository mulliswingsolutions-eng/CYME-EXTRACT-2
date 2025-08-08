"""
Bus.py
------
Bus extraction utilities for CYME text export (XML content).

Public API:
    extract_bus_data(path: str | Path) -> pandas.DataFrame

Columns returned (where available):
    - NodeID
    - X
    - Y
    - BusWidth
    - TagText
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


def extract_bus_data(path: str | Path) -> pd.DataFrame:
    """
    Parse bus / node information from a CYME text export.

    Parameters
    ----------
    path : str | Path

    Returns
    -------
    pandas.DataFrame
        One row per Node (i.e., "bus").
    """
    p = Path(path)
    tree = ET.parse(p)
    root = tree.getroot()

    network = _get_network(root)
    if network is None:
        return pd.DataFrame(columns=["NodeID", "X", "Y", "BusWidth", "TagText"])

    nodes = network.find("Nodes")
    if nodes is None:
        return pd.DataFrame(columns=["NodeID", "X", "Y", "BusWidth", "TagText"])

    rows: List[Dict[str, Any]] = []

    for node in nodes.findall("Node"):
        node_id = _first_text(node.find("NodeID"))

        # Coordinates live under Connectors/Point
        x = y = None
        connectors = node.find("Connectors")
        if connectors is not None:
            pt = connectors.find("Point")
            if pt is not None:
                x = _float_or_none(_first_text(pt.find("X")))
                y = _float_or_none(_first_text(pt.find("Y")))

        # If a BusDisplay is present, width can be helpful for visuals
        bus_width = None
        bus_disp = node.find("BusDisplay")
        if bus_disp is not None:
            bus_width = _float_or_none(_first_text(bus_disp.find("Width")))

        # Optional tag text (often contains voltage placeholders)
        tag_text = ""
        tag = node.find("Tag")
        if tag is not None:
            tag_text = _first_text(tag.find("Text"))

        rows.append(
            {
                "NodeID": node_id,
                "X": x,
                "Y": y,
                "BusWidth": bus_width,
                "TagText": tag_text,
            }
        )

    # Sort by NodeID when possible (numeric IDs numerically, otherwise lexicographically)
    def _sort_key(v: str):
        try:
            return (0, int(v))
        except Exception:
            return (1, v)

    rows.sort(key=lambda r: _sort_key(r["NodeID"]))

    return pd.DataFrame(rows, columns=["NodeID", "X", "Y", "BusWidth", "TagText"])
