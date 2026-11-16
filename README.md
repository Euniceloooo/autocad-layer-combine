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
