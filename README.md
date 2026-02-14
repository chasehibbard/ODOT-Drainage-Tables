# ODOT Drainage Tables Generator

Converts raw flex table exports from OpenRoads Designer into formatted ODOT
(Oklahoma Department of Transportation) drainage tables for plan sheets.

## What It Does

- Reads raw `.xlsx` exports from OpenRoads Designer
- Generates two formatted output tables:
  - **Drainage Area Summary Table**
  - **Inlet Design Record**
- Outputs a single `.xlsx` file matching ODOT formatting standards

## Requirements

- Python 3.8 or newer
- Install dependencies: `pip install -r requirements.txt`

## How to Use

1. Export your flex tables from OpenRoads Designer as `.xlsx` files
2. Run the program:
   ```
   python src/main.py
   ```
3. Use the file picker to select your input files
4. Choose where to save the output
5. Done — your formatted ODOT drainage table is ready

## Input Files Expected

These are the raw OpenRoads flex table exports:

| File | Used For |
|------|----------|
| `OKDOT_Drainage Areas Summary Table.xlsx` | DA Summary Table |
| `OKDOT_Inlets for Storm Sewer Design Record.xlsx` | Inlet Design Record |

## Output

A single Excel workbook with two sheets:
- **Drainage Area Summary** — DA designations, areas, Tc, C values, rainfall intensity, and peak flow
- **Inlets** — Structure numbers, locations, inlet types, hydraulic data, spread, bypass info
