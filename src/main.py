# src/main.py
from __future__ import annotations
from pathlib import Path
import pandas as pd

# Local imports (same folder)
from Modules.General import get_general
from Modules.Bus import extract_bus_data
from Modules.Voltage_Source import write_voltage_source_sheet
from Modules.Load import write_load_sheet
from Modules.Line import write_line_sheet  # <-- NEW

# ===== Paths (adjust as needed) =====
INPUT_PATH = Path(__file__).parent.parent / "Examples/Example-13bus-modified.txt"
OUTPUT_PATH = Path(__file__).parent.parent / "Outputs/CYME_Extract_13Bus.xlsx"


def main():
    in_path = INPUT_PATH.resolve()
    out_path = OUTPUT_PATH.resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    if not in_path.exists():
        raise FileNotFoundError(f"Input not found: {in_path}")

    # Parse data needed for General/Bus sheets
    general_rows = get_general(in_path)
    bus_rows = extract_bus_data(in_path)
    df_bus = pd.DataFrame(bus_rows)

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as xw:
        # General
        pd.DataFrame(general_rows).to_excel(xw, sheet_name="General", index=False, header=False)

        # Bus
        if df_bus.empty:
            df_bus = pd.DataFrame(columns=["Bus", "Base Voltage (V)", "Initial Vmag", "Unit", "Angle", "Type"])
        df_bus.to_excel(xw, sheet_name="Bus", index=False)

        # Voltage Source
        write_voltage_source_sheet(xw, in_path)

        # Load
        write_load_sheet(xw, in_path)

        # Line  <-- NEW
        write_line_sheet(xw, in_path)

    print(f"Wrote: {out_path}")


if __name__ == "__main__":
    main()
