"""
main.py
-------
End-to-end runner: parse a CYME text export and write three sheets to an XLSX:
    - General
    - Bus
    - Voltage Source
"""

from __future__ import annotations

from pathlib import Path
import pandas as pd

# ===== USER-CONFIGURABLE PATHS =====
INPUT_PATH = Path(__file__).parent.parent / "Examples" / "Example-13bus-modified.txt"
OUTPUT_PATH = Path(__file__).parent.parent / "Outputs" / "CYME_Extract_13Bus.xlsx"
# ===================================

# Local imports
from Modules.General import get_general
from Modules.Bus import extract_bus_data
from Modules.Voltage_Source import extract_voltage_source_data


def _fmt_num(x):
    """Pretty number formatter: 60 -> '60', 60.5 -> '60.5'."""
    try:
        f = float(x)
    except Exception:
        return x
    if f.is_integer():
        return str(int(f))
    return str(f)


def main():
    input_path = INPUT_PATH.resolve()
    output_path = OUTPUT_PATH.resolve()

    if not input_path.exists():
        print(f"Input not found: {input_path}")
        return

    # Parse
    rows_general = get_general(input_path)
    df_bus = pd.DataFrame(extract_bus_data(input_path))
    df_vs = extract_voltage_source_data(input_path)  # already a DataFrame in template format

    # Print to terminal
    print("=== General ===")
    for field, value in rows_general:
        print(f"{field:<24} {_fmt_num(value)}")

    print("\n=== Bus ===")
    if df_bus.empty:
        print("(no bus rows)")
    else:
        print(df_bus.to_string(index=False))

    print("\n=== Voltage Source ===")
    if df_vs.empty:
        print("(no voltage source rows)")
    else:
        print(df_vs.to_string(index=False, header=False))

    # Write Excel
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as xw:
        # General sheet (no headers)
        pd.DataFrame(rows_general).to_excel(
            xw, index=False, header=False, sheet_name="General"
        )
        # Bus sheet
        (df_bus if not df_bus.empty else pd.DataFrame(
            columns=["NodeID", "X", "Y", "BusWidth", "TagText"]
        )).to_excel(xw, index=False, sheet_name="Bus")
        # Voltage Source sheet (already in correct template layout)
        (df_vs if not df_vs.empty else pd.DataFrame()).to_excel(
            xw, index=False, header=False, sheet_name="Voltage Source"
        )

    print(f"\nWrote sheets to: {output_path}")


if __name__ == "__main__":
    main()
