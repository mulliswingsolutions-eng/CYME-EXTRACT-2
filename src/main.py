# To change:
# 1. Remove 12.47 Feeders as sources (done)
# 2. Check out constant current loads (done)
# 3. Double check load values (Done)
# 5. Number of taps on transformers (Done)

# src/main.py
from __future__ import annotations
from pathlib import Path
import pandas as pd  # needed for ExcelWriter (modules expect xw.book / xw.sheets)

# Local imports
from Modules.Explorer import write_overview_sheet
from Modules.General import write_general_sheet
from Modules.Pins import write_pins_sheet
from Modules.Bus import write_bus_sheet
from Modules.Voltage_Source import write_voltage_source_sheet
from Modules.Load import write_load_sheet
from Modules.Line import write_line_sheet
from Modules.Transformer import write_transformer_sheet
from Modules.Shunt import write_shunt_sheet
from Modules.Switch import write_switch_sheet

# ===== Paths (adjust as needed) =====

#INPUT_PATH = Path(__file__).parent.parent / "Examples/Example-13bus-modified.txt"
#OUTPUT_PATH = Path(__file__).parent.parent / "Outputs/CYME_Extract_13Bus.xlsx"

#INPUT_PATH = Path(__file__).parent.parent / "Examples/Example-4bus.txt"
#OUTPUT_PATH = Path(__file__).parent.parent / "Outputs/CYME_Extract_4Bus.xlsx"

#INPUT_PATH = Path(__file__).parent.parent / "Examples/Saint-John-CYME.txt"
#OUTPUT_PATH = Path(__file__).parent.parent / "Outputs/CYME_Extract_Saint-John.xlsx"

INPUT_PATH = Path(__file__).parent.parent / "Examples/UNB Feeder.txt"
OUTPUT_PATH = Path(__file__).parent.parent / "Outputs/CYME_Extract_UNB.xlsx"

# ====================================


def main():
    in_path = INPUT_PATH.resolve()
    out_path = OUTPUT_PATH.resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    if not in_path.exists():
        raise FileNotFoundError(f"Input not found: {in_path}")

    # Create workbook and let each module render its own sheet
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as xw:
        write_general_sheet(xw, in_path)
        #write_overview_sheet(xw, in_path)
        write_pins_sheet(xw, in_path)
        write_bus_sheet(xw, in_path)
        write_voltage_source_sheet(xw, in_path)
        write_load_sheet(xw, in_path)
        write_line_sheet(xw, in_path)
        write_transformer_sheet(xw, in_path)
        write_shunt_sheet(xw, in_path)
        write_switch_sheet(xw, in_path)

    print(f"Wrote: {out_path}")


if __name__ == "__main__":
    main()
