"""
Combine selected layers from multiple AutoCAD DWGs into one drawing,
renaming layers with each file's TLC suffix, and mark the origin with a 1 m cross.
Requires AutoCAD installed. Run from the project .venv only.
"""

from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path

try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
    from win32com.client import VARIANT  # type: ignore
except ImportError:
    print(
        "pywin32 is not installed in this Python environment.\n"
        "Activate the project .venv and run: pip install -r requirements.txt"
    )
    sys.exit(1)


SUFFIX_RE = re.compile(r"^(TLC\d+)$", re.IGNORECASE)
ORIGIN_LAYER = "ORIGIN_MARK"
ACAD_COLOR_RED = 1

# #region agent log
_DEBUG_LOG_PATH = Path(__file__).resolve().parent / ".cursor" / "debug-8178ab.log"


def _agent_log(hypothesis_id: str, location: str, message: str, data: dict | None = None) -> None:
    payload = {
        "sessionId": "8178ab",
        "runId": "post-fix-cross",
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": data or {},
        "timestamp": int(time.time() * 1000),
    }
    try:
        _DEBUG_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with _DEBUG_LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(json.dumps(payload, default=str) + "\n")
    except Exception:
        pass
    print(f"[debug:{hypothesis_id}] {message}")


# #endregion


def _acad_point(x: float, y: float, z: float = 0.0):
    """AutoCAD COM points must be VT_ARRAY|VT_R8, not plain Python tuples."""
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def parse_suffix(dwg_path: Path) -> str:
    """Last _-separated token of the stem, preferably TLC## without leading underscore."""
    parts = dwg_path.stem.split("_")
    if not parts:
        return dwg_path.stem
    token = parts[-1]
    m = SUFFIX_RE.match(token)
    if m:
        return m.group(1).upper()
    return token


def prompt_folder() -> Path:
    while True:
        raw = input("Enter folder path containing DWG files: ").strip().strip('"')
        folder = Path(raw)
        if folder.is_dir():
            return folder.resolve()
        print(f"Not a valid folder: {folder}")


def list_dwgs(folder: Path) -> list[Path]:
    files = sorted(folder.glob("*.dwg")) + sorted(folder.glob("*.DWG"))
    seen: set[str] = set()
    unique: list[Path] = []
    for f in files:
        key = str(f).lower()
        if key not in seen:
            seen.add(key)
            unique.append(f)
    return unique


def get_acad():
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception:
        acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True
    return acad


def _is_document(obj) -> bool:
    if obj is None:
        return False
    username = getattr(obj, "_username_", "") or ""
    if username in ("Open", "Documents", "Application"):
        return False
    try:
        _ = obj.Layers
        _ = obj.Name
        return True
    except Exception:
        return False


def open_dwg(acad, path: Path | str, read_only: bool = True):
    """
    Open a DWG and return a real Document.

    Late-bound Documents.Open often does not return a Document (win32com leaves
    the Open method object). Prefer ActiveDocument after Open.
    """
    path_str = str(path)
    opened = None
    try:
        opened = acad.Documents.Open(path_str, read_only)
    except Exception:
        opened = acad.Documents.Open(path_str)

    opened_ok = _is_document(opened)
    # #region agent log
    _agent_log(
        "H1",
        "main.py:open_dwg",
        "Documents.Open returned",
        {
            "path": path_str,
            "openedOk": opened_ok,
            "openedUsername": getattr(opened, "_username_", None),
            "openedType": type(opened).__name__,
        },
    )
    # #endregion
    if opened_ok:
        return opened

    doc = acad.ActiveDocument
    active_ok = _is_document(doc)
    # #region agent log
    _agent_log(
        "H2",
        "main.py:open_dwg",
        "ActiveDocument fallback",
        {
            "path": path_str,
            "activeOk": active_ok,
            "activeName": str(getattr(doc, "Name", "?")) if active_ok else None,
            "activeUsername": getattr(doc, "_username_", None),
        },
    )
    # #endregion
    if not active_ok:
        raise RuntimeError(
            f"Failed to get Document after Open: {path_str} "
            f"(opened type={type(opened).__name__}, "
            f"username={getattr(opened, '_username_', None)})"
        )
    return doc


def collect_layer_names(acad, dwg_paths: list[Path]) -> list[str]:
    """Union of layer names from all source drawings."""
    names: set[str] = set()
    # Step-by-step: scan first drawing for the layer picker (faster).
    # Kept-layer matching later still works across all files.
    scan_paths = dwg_paths[:1] if dwg_paths else []
    for path in scan_paths:
        doc = open_dwg(acad, path, read_only=True)
        # #region agent log
        _agent_log(
            "H3",
            "main.py:collect_layer_names",
            "before Layers.Count",
            {
                "path": str(path),
                "docOk": _is_document(doc),
                "docName": str(getattr(doc, "Name", "?")),
                "docUsername": getattr(doc, "_username_", None),
            },
        )
        # #endregion
        try:
            for i in range(doc.Layers.Count):
                layer = doc.Layers.Item(i)
                name = str(layer.Name)
                if name.upper() in ("0", "DEFPOINTS"):
                    continue
                names.add(name)
            # #region agent log
            _agent_log(
                "H3",
                "main.py:collect_layer_names",
                "layer scan ok",
                {"layerCount": len(names)},
            )
            # #endregion
        finally:
            doc.Close(False)
    return sorted(names, key=str.lower)


def prompt_layers_to_keep(available: list[str]) -> list[str]:
    print("\nAvailable layers:")
    for name in available:
        print(f"  - {name}")
    print(
        "\nEnter layer name(s) to keep (comma-separated). "
        "Match is case-insensitive."
    )
    while True:
        raw = input("Layers to keep: ").strip()
        if not raw:
            print("Enter at least one layer name.")
            continue
        requested = [p.strip() for p in raw.split(",") if p.strip()]
        lower_map = {n.lower(): n for n in available}
        kept: list[str] = []
        missing: list[str] = []
        for req in requested:
            hit = lower_map.get(req.lower())
            if hit:
                if hit not in kept:
                    kept.append(hit)
            else:
                missing.append(req)
        if missing:
            print(f"Warning — not found (will skip): {', '.join(missing)}")
        if not kept:
            print("No matching layers. Try again.")
            continue
        print(f"Keeping: {', '.join(kept)}")
        return kept


def prompt_origin() -> tuple[float, float]:
    print(
        "\nOrigin mark (1 m cross). Enter Easting (E) and Northing (N) in METRES. "
        "Z will be 0. CAD units should already be metres."
    )
    while True:
        try:
            e = float(input("Easting E (metres): ").strip())
            n = float(input("Northing N (metres): ").strip())
            return e, n
        except ValueError:
            print("Please enter numeric values in metres.")


def prompt_output_path(folder: Path) -> Path:
    default = folder / "combined_layers.dwg"
    raw = input(f"Output DWG path [{default}]: ").strip().strip('"')
    if not raw:
        return default
    path = Path(raw)
    if path.suffix.lower() != ".dwg":
        path = path.with_suffix(".dwg")
    return path.resolve()


def ensure_layer(doc, name: str, color: int | None = None):
    try:
        layer = doc.Layers.Item(name)
    except Exception:
        layer = doc.Layers.Add(name)
    if color is not None:
        try:
            layer.Color = color
        except Exception:
            pass
    return layer


def _as_entity_list(result) -> list:
    if result is None:
        return []
    if hasattr(result, "Count"):
        return [result.Item(j) for j in range(result.Count)]
    if isinstance(result, (list, tuple)):
        return list(result)
    return [result]


def copy_kept_entities(
    acad,
    source_path: Path,
    target_doc,
    kept_layers: list[str],
    suffix: str,
) -> int:
    """Copy modelspace entities on kept layers into target; rename to layer_suffix."""
    kept_lower = {k.lower() for k in kept_layers}
    source = open_dwg(acad, source_path, read_only=True)
    copied = 0
    try:
        layer_map: dict[str, str] = {}
        for i in range(source.Layers.Count):
            src_layer = source.Layers.Item(i)
            src_name = str(src_layer.Name)
            if src_name.lower() not in kept_lower:
                continue
            dst_name = f"{src_name}_{suffix}"
            layer_map[src_name] = dst_name
            ensure_layer(target_doc, dst_name)

        if not layer_map:
            return 0

        ms = source.ModelSpace
        to_copy = []
        for i in range(ms.Count):
            ent = ms.Item(i)
            try:
                layer_name = str(ent.Layer)
            except Exception:
                continue
            if layer_name in layer_map:
                to_copy.append(ent)

        if not to_copy:
            return 0

        new_objs: list = []
        try:
            obj_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, to_copy)
            new_objs = _as_entity_list(
                source.CopyObjects(obj_array, target_doc.ModelSpace)
            )
        except Exception:
            for ent in to_copy:
                try:
                    one = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [ent])
                    new_objs.extend(
                        _as_entity_list(
                            source.CopyObjects(one, target_doc.ModelSpace)
                        )
                    )
                except Exception as ex:
                    print(f"  Skip entity copy error: {ex}")

        for idx, new_ent in enumerate(new_objs):
            try:
                if idx < len(to_copy):
                    src_layer = str(to_copy[idx].Layer)
                    dst = layer_map.get(src_layer)
                    if dst:
                        new_ent.Layer = dst
                copied += 1
            except Exception:
                copied += 1

        if not new_objs and to_copy:
            tms = target_doc.ModelSpace
            for i in range(tms.Count):
                ent = tms.Item(i)
                try:
                    ln = str(ent.Layer)
                    if ln in layer_map:
                        ent.Layer = layer_map[ln]
                        copied += 1
                except Exception:
                    pass
    finally:
        source.Close(False)

    return copied


def draw_origin_cross(doc, easting: float, northing: float) -> None:
    """1 m cross at (E, N, 0) plus a text label keeping E/N values visible."""
    # #region agent log
    _agent_log(
        "A",
        "main.py:draw_origin_cross",
        "before AddLine",
        {
            "easting": easting,
            "northing": northing,
            "eType": type(easting).__name__,
            "nType": type(northing).__name__,
            "docName": str(getattr(doc, "Name", "?")),
        },
    )
    # #endregion
    try:
        doc.Activate()
    except Exception:
        pass
    ensure_layer(doc, ORIGIN_LAYER, color=ACAD_COLOR_RED)
    ms = doc.ModelSpace
    half = 0.5
    p1 = _acad_point(easting - half, northing, 0.0)
    p2 = _acad_point(easting + half, northing, 0.0)
    p3 = _acad_point(easting, northing - half, 0.0)
    p4 = _acad_point(easting, northing + half, 0.0)
    try:
        h = ms.AddLine(p1, p2)
        h.Layer = ORIGIN_LAYER
        v = ms.AddLine(p3, p4)
        v.Layer = ORIGIN_LAYER
        # #region agent log
        _agent_log(
            "A",
            "main.py:draw_origin_cross",
            "AddLine ok with VARIANT points",
            {"half": half},
        )
        # #endregion
    except Exception as ex:
        # #region agent log
        _agent_log(
            "A",
            "main.py:draw_origin_cross",
            "AddLine failed",
            {"error": repr(ex)},
        )
        # #endregion
        raise

    # B: keep E/N as readable text near the cross (offset slightly so it doesn't sit on the lines)
    label = f"ORIGIN  E={easting}  N={northing}  Z=0"
    text_height = 0.15  # metres
    text_pt = _acad_point(easting + 0.15, northing + 0.15, 0.0)
    try:
        txt = ms.AddText(label, text_pt, text_height)
        txt.Layer = ORIGIN_LAYER
        # #region agent log
        _agent_log(
            "B",
            "main.py:draw_origin_cross",
            "AddText ok",
            {"label": label, "height": text_height},
        )
        # #endregion
    except Exception as ex:
        # #region agent log
        _agent_log(
            "B",
            "main.py:draw_origin_cross",
            "AddText failed",
            {"error": repr(ex)},
        )
        # #endregion
        raise


def main() -> int:
    print("AutoCAD layer combine")
    print("Uses AutoCAD COM. Packages must be installed in project .venv only.\n")
    # #region agent log
    _agent_log(
        "H0",
        "main.py:main",
        "run start A+B build",
        {"script": str(Path(__file__).resolve()), "py": sys.version.split()[0]},
    )
    # #endregion

    folder = prompt_folder()
    dwgs = list_dwgs(folder)
    if not dwgs:
        print(f"No DWG files found in {folder}")
        return 1

    print(f"\nFound {len(dwgs)} DWG(s):")
    file_suffixes: list[tuple[Path, str]] = []
    for p in dwgs:
        suffix = parse_suffix(p)
        file_suffixes.append((p, suffix))
        print(f"  {p.name}  →  suffix {suffix}")

    print("\nStarting AutoCAD (if not already running)...")
    acad = get_acad()

    print("Scanning layers (opening drawings read-only)...")
    available = collect_layer_names(acad, dwgs)
    if not available:
        print("No layers found in the drawings.")
        return 1

    kept = prompt_layers_to_keep(available)
    easting, northing = prompt_origin()
    out_path = prompt_output_path(folder)

    # Documents.Add can have the same late-binding return quirk as Open
    added = acad.Documents.Add()
    target = added if _is_document(added) else acad.ActiveDocument
    if not _is_document(target):
        raise RuntimeError("Failed to create target drawing")
    total_copied = 0
    per_file: list[tuple[str, str, int]] = []

    for path, suffix in file_suffixes:
        print(f"\nProcessing {path.name} (suffix {suffix})...")
        n = copy_kept_entities(acad, path, target, kept, suffix)
        per_file.append((path.name, suffix, n))
        total_copied += n
        print(f"  Copied {n} entity(ies)")

    print("\nDrawing 1 m origin cross...")
    # Re-resolve target in case ActiveDocument changed during copy loop
    try:
        if str(acad.ActiveDocument.Name) == str(target.Name):
            target = acad.ActiveDocument
    except Exception:
        pass
    draw_origin_cross(target, easting, northing)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        try:
            out_path.unlink()
        except OSError as ex:
            print(f"Could not overwrite existing file: {ex}")
            return 1

    target.SaveAs(str(out_path))
    print(f"\nSaved: {out_path}")

    print("\n--- Summary ---")
    print(f"Files processed: {len(file_suffixes)}")
    print(f"Layers kept: {', '.join(kept)}")
    print(f"Origin (E, N, Z): ({easting}, {northing}, 0) metres")
    for name, suffix, n in per_file:
        print(f"  {name} [{suffix}]: {n} entities → layers *_{suffix}")
    print(f"Total entities copied: {total_copied}")
    print(f"Output: {out_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nCancelled.")
        raise SystemExit(130)
