# autocad-layer-combine

Combine selected layers from multiple AutoCAD DWG files into one drawing.
Each source file’s trailing token (e.g. `TLC01`) is appended to kept layer names
as `{layer}_{suffix}`. A 1 m origin cross is drawn at a user-supplied E/N (metres).

## Requirements

- Windows with **AutoCAD** installed (COM automation)
- Project **`.venv`** only — do not install packages into system Python

## Setup

Install into the project `.venv` only (not system Python):

```bat
cd c:\Users\Eunice.Lo\github\autocad-layer-combine
install_deps.bat
```

Or manually:

```bat
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
```

## Run

With AutoCAD available on the machine:

```bat
.venv\Scripts\activate.bat
python main.py
```

You will be prompted for:

1. Folder containing DWGs (e.g. `..._CHAINAGES_TLC01.dwg`)
2. Layer name(s) to keep (comma-separated)
3. Origin Easting / Northing in **metres** (Z = 0) — skipped if `origin.local.json` exists
4. Output path (default: `combined_layers.dwg` in that folder)

### Local origin (not on GitHub)

Copy the template and fill in your HK Grid origin (metres):

```bat
copy origin.example.json origin.local.json
```

Edit `origin.local.json` with your `easting`, `northing`, and `elevation`. That file is **gitignored** so real E/N are not pushed to GitHub. When it is present, `main.py` uses it automatically and does not ask for E/N.

## Notes

- Suffix is stored as `TLC01` (no leading `_`); rename uses `f"{layer}_{suffix}"`.
- Origin mark: 1 m horizontal + 1 m vertical cross centered on `(E, N, 0)`, plus a text label `ORIGIN E=... N=... Z=0` on layer `ORIGIN_MARK`.

## 3D polyline geometry export

Export each `AcDb3dPolyline` in one DWG to a CSV named by **layer**. Finds polylines in model space and inside **block inserts** (vertices transformed to WCS).

```bat
RUN_EXTRACT_GEOMETRY.bat
```

Or:

```bat
.venv\Scripts\activate.bat
python extract_3d_polyline_geometry.py
```

Requires `origin.local.json` (copy from `origin.example.json`) for the adaptive offset.

Prompts:

1. Path to a single `.dwg`
2. Output folder (default: same folder as the DWG)

For each layer with exactly one 3D polyline, writes **two** CSVs:

| File | Coordinates |
|------|-------------|
| `{Layer}_Geometry.csv` | HK Grid absolute (metres) |
| `{Layer}_Geometry for adaptive.csv` | Local: `x−E0`, `y−N0`, `z−Z0` from `origin.local.json` |

```csv
vertex,x,y,z
1,2.9990,0.0000,12.3000
2,...
```

Values are rounded to **4 decimal places** (0.1 mm). Use the adaptive CSV for Revit adaptive-family Reference Points (near origin).

**One 3D polyline per layer** is expected. If a layer has more than one (including from blocks), the tool prints a **warning** and **skips** that layer (no CSV).
