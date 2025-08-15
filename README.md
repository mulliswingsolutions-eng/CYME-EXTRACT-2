# Create a plain-text README file for the CYME Extract project
readme_text = """CYME EXTRACT
==============

Parse CYME text/XML exports and generate a clean Excel workbook with sheets:
Bus, Line, Load, Transformer, Voltage Source, Switch.

Quick Start
-----------
1) Create & activate a venv
   Windows:   python -m venv venv && .\\venv\\Scripts\\activate
   macOS/Linux: python -m venv venv && source venv/bin/activate

2) Install deps
   pip install -r requirements.txt

3) Run
   python -m src.main -i data/Example-13bus-modified.txt -o out/CYME_Extract.xlsx
   # or use default output path:
   python -m src.main -i data/Example-13bus-modified.txt

requirements.txt
----------------
xlsxwriter>=3.2.0

What it Extracts
----------------
- Voltage Source: Detects Short-Circuit vs Sequential models. Uses SC MVA when present
  (replaces nominal 200000 only if no capacity is provided). Angle_a is 0 unless an offset exists.
- Transformer: Reads kV/kVA/Z%/X/R and connection from TransformerDB only (no hardcoded defaults).
  Computes X (pu) from Z% & X/R; splits R into RW1/RW2.
- Load: Binds to FromNodeID. Converts KW/PF, KVA/PF, KVA/kVAr, KW/kVAr. Classifies 1φ / 2φ / 3φ
  from CustomerLoadValue phases.
- Line / Cable: Uses OverheadLineDB (sequence) or OverheadLineUnbalancedDB (full ABC) by LineID.
  Converts per-km → per-mile; phase-aware bus labels.
- Switching Devices: Unified list for Switch, Sectionalizer, Breaker, Fuse.
  Per-phase status from ClosedPhase or NormalStatus.

Excel Layout Rules
------------------
- Blocks start at row 11; "Go to Type List" links jump back to A1.
- Exactly one blank row between blocks; otherwise no empty rows.
- Columns match planning template.

CLI
---
python -m src.main -i <path_to_cyme_export.txt> [-o <path_to_output.xlsx>]

- -i, --input  : CYME text/XML export file.
- -o, --output : Target Excel (default: ./out/CYME_Extract.xlsx).

Repo Layout
-----------
src/
  main.py                (orchestrates parsing + workbook writing)
  Modules/
    Bus.py
    Line.py
    Load.py
    Transformer.py
    Voltage_Source.py
    Switch.py
tests/
  test_*.py
requirements.txt

Tips
----
- If a line has <LineID>…</LineID>, impedances are looked up by EquipmentID in the matching *LineDB.
- Single-phase loads (e.g., Phase=A only) appear in the Single-Phase block even if the Section Phase is ABC.
- Devices that are graphically present but normally open still appear; per-phase Status reflects ClosedPhase/NormalStatus.
"""