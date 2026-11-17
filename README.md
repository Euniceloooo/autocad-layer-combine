# autocad-layer-combine

Windows tools that use **AutoCAD COM** (via project `.venv` + `pywin32`):

1. **Layer combine** — merge selected layers from many DWGs, rename by file suffix, mark origin.
2. **3D polyline geometry export** — CSV per layer (HK Grid + local-offset for Revit adaptive families).

## Requirements

- Windows with **AutoCAD** installed
- Project **`.venv` only** — do not install packages into system Python

## Setup

```bat
cd c:\Users\Eunice.Lo\github\autocad-layer-combine
install_deps.bat
```

Or:

```bat
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
```

## Local origin (not on GitHub)

Both tools can use a fixed HK Grid origin from a **gitignored** file:

```bat
copy origin.example.json origin.local.json
```

Edit `origin.local.json`:

```json
{
  "easting": 0.0,
  "northing": 0.0,
  "elevation": 0.0
}
```

Real E/N stay off GitHub. Geometry export **requires** this file (for the adaptive CSV). Layer combine uses it to skip typing E/N when present.

---

## 1. Layer combine

```bat
RUN_MAIN.bat
```

Or: `.venv\Scripts\python.exe -u main.py`

Prompts:

1. Folder containing DWGs (e.g. `..._CHAINAGES_TLC01.dwg`)
2. Layer name(s) to keep (comma-separated)
3. Origin E/N in metres — **skipped** if `origin.local.json` exists
4. Output path (default: `combined_layers.dwg` in that folder)

Behaviour:

- Suffix from the last `_` token of each filename (e.g. `TLC01`, `TMR-EB`)
- Kept layers become `{layer}_{suffix}`
- Draws a 1 m origin cross + text `ORIGIN E=... N=... Z=0` on layer `ORIGIN_MARK`

---

## 2. 3D polyline geometry export

```bat
RUN_EXTRACT_GEOMETRY.bat
```

Or: `.venv\Scripts\python.exe -u extract_3d_polyline_geometry.py`

Requires `origin.local.json` first (see above).

Prompts:

1. Path to a single `.dwg`
2. Output folder (default: same folder as the DWG)

Finds `AcDb3dPolyline` in **model space** and inside **block inserts** (vertices to WCS).

**One 3D polyline per layer** is expected. If a layer has more than one, the tool **warns and skips** that layer.

For each OK layer, writes **two** CSVs (coordinates in metres, **4 decimal places** = 0.1 mm):

| File | Coordinates |
|------|-------------|
| `{Layer}_Geometry.csv` | HK Grid absolute |
| `{Layer}_Geometry for adaptive.csv` | Local: `x−E0`, `y−N0`, `z−Z0` from `origin.local.json` |

Example adaptive row (near origin):

```csv
vertex,x,y,z
1,2.9990,0.0000,12.3000
2,...
```

Use `*_Geometry for adaptive.csv` when placing Reference Points in a Revit adaptive family (values near 0, not ~820000).
