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
3. Origin Easting / Northing in **metres** (Z = 0)
4. Output path (default: `combined_layers.dwg` in that folder)

## Notes

- Suffix is stored as `TLC01` (no leading `_`); rename uses `f"{layer}_{suffix}"`.
- Origin mark: 1 m horizontal + 1 m vertical cross centered on `(E, N, 0)`, plus a text label `ORIGIN E=... N=... Z=0` on layer `ORIGIN_MARK`.

## 3D polyline geometry export

Export each `AcDb3dPolyline` in one DWG to its own CSV (by AutoCAD object Handle).

```bat
RUN_EXTRACT_GEOMETRY.bat
```

Or:

```bat
.venv\Scripts\activate.bat
python extract_3d_polyline_geometry.py
```

Prompts:

1. Path to a single `.dwg`
2. Output folder (default: same folder as the DWG)

For each 3D polyline, writes `{Handle}_Geometry.csv`:

```csv
vertex,x,y,z
1,821704.4910,824308.5050,12.3000
2,...
```

Coordinates are **HK Grid metres**, rounded to **4 decimal places** (0.1 mm) to keep survey accuracy while trimming CAD float noise.

Multiple 3D polylines in one drawing are processed one by one (one CSV each).
