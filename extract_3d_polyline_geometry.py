"""
Extract 3D polyline vertices from one AutoCAD DWG into per-polyline CSV files.

Each AcDb3dPolyline becomes `{Handle}_Geometry.csv` with columns:
  vertex,x,y,z

Requires AutoCAD installed. Run from the project .venv only.
"""

from __future__ import annotations

import csv
import sys
import time
from pathlib import Path

try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
except ImportError:
    print(
        "pywin32 is not installed in this Python environment.\n"
        "Activate the project .venv and run: pip install -r requirements.txt"
    )
    sys.exit(1)


_RPC_E_CALL_REJECTED = -2147418111
_RPC_E_SERVERCALL_RETRYLATER = -2147417846
_COM_BUSY = {_RPC_E_CALL_REJECTED, _RPC_E_SERVERCALL_RETRYLATER}

# AutoCAD entity ObjectName for 3D polylines
OBJ_3DPOLY = "AcDb3dPolyline"


def _com_hresult(ex: BaseException) -> int | None:
    try:
        return int(ex.args[0])
    except Exception:
        return None


def _com_retry(fn, *, retries: int = 50, delay: float = 0.2, label: str = "com"):
    last: BaseException | None = None
    for attempt in range(1, retries + 1):
        try:
            return fn()
        except Exception as ex:
            last = ex
            if _com_hresult(ex) not in _COM_BUSY:
                raise
            if attempt == 1 or attempt % 10 == 0:
                print(f"  (AutoCAD busy — retry {attempt}/{retries}: {label})")
            try:
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass
            time.sleep(delay)
    assert last is not None
    raise last


def _close_doc(doc) -> None:
    def _do():
        doc.Close(False)

    try:
        _com_retry(_do, label="Close")
    except Exception as ex:
        print(f"  Warning: Close failed: {ex}")


class _AcadComMessageFilter:
    def HandleInComingCall(self, dwCallType, hTaskCaller, dwTickCount, lpInterfaceInfo):
        return 1

    def RetryRejectedCall(self, hTaskCallee, dwTickCount, dwRejectType):
        if dwTickCount < 30000:
            return 250
        return -1

    def MessagePending(self, hTaskCallee, dwTickCount, dwPendingType):
        return 2


def get_acad():
    try:
        pythoncom.CoRegisterMessageFilter(_AcadComMessageFilter())
    except Exception:
        pass
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception:
        acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True

    def _ping():
        _ = acad.Name
        return acad

    return _com_retry(_ping, label="get_acad", retries=60, delay=0.5)


def _is_document(obj) -> bool:
    if obj is None:
        return False
    username = getattr(obj, "_username_", "") or ""
    if username in ("Open", "Documents", "Application"):
        return False
    try:
        _ = obj.Name
        return True
    except Exception as ex:
        if _com_hresult(ex) in _COM_BUSY:
            raise
        return False


def open_dwg(acad, path: Path | str, read_only: bool = True):
    path_str = str(Path(path).resolve())
    expected = Path(path_str).name.lower()

    def _find_open():
        docs = acad.Documents
        for i in range(int(docs.Count)):
            d = docs.Item(i)
            try:
                name = str(d.Name).lower()
                full = str(getattr(d, "FullName", "")).lower()
            except Exception:
                continue
            if name == expected or full.replace("/", "\\") == path_str.lower():
                return d
        return None

    def _existing_doc():
        d = _find_open()
        if d is None:
            return None
        return d if _is_document(d) else None

    existing = _com_retry(_existing_doc, label="findOpen")
    if existing is not None:
        try:
            _com_retry(lambda: existing.Activate(), label="Activate")
        except Exception:
            pass
        return existing

    def _do_open():
        try:
            return acad.Documents.Open(path_str, read_only)
        except Exception:
            return acad.Documents.Open(path_str)

    opened = _com_retry(_do_open, label="Open")
    try:
        if _is_document(opened):
            return opened
    except Exception:
        pass

    doc = None
    for attempt in range(1, 51):
        try:
            doc = acad.ActiveDocument
            if _is_document(doc):
                name = str(doc.Name).lower()
                full = str(getattr(doc, "FullName", "")).lower().replace("/", "\\")
                if name == expected or full == path_str.lower():
                    break
                found = _find_open()
                if found is not None and _is_document(found):
                    doc = found
                    break
        except Exception as ex:
            if _com_hresult(ex) not in _COM_BUSY:
                raise
            if attempt == 1 or attempt % 10 == 0:
                print(f"  (AutoCAD busy — retry {attempt}/50: ActiveDocument)")
            try:
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass
            time.sleep(0.2)
            continue
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass
        time.sleep(0.2)

    if not _is_document(doc):
        raise RuntimeError(f"Failed to get Document after Open: {path_str}")
    return doc


def prompt_dwg_path() -> Path:
    while True:
        raw = input("Enter DWG file path: ").strip().strip('"')
        path = Path(raw)
        if path.is_file() and path.suffix.lower() == ".dwg":
            return path.resolve()
        print(f"Not a valid .dwg file: {path}")


def prompt_output_folder(default: Path) -> Path:
    raw = input(f"Output folder for CSV files [{default}]: ").strip().strip('"')
    folder = Path(raw) if raw else default
    folder.mkdir(parents=True, exist_ok=True)
    return folder.resolve()


def _coords_to_vertices(coords) -> list[tuple[float, float, float]]:
    """Flatten AutoCAD Coordinates (x,y,z,x,y,z,...) into vertex tuples."""
    if coords is None:
        return []
    if hasattr(coords, "tolist"):
        vals = list(coords.tolist())
    else:
        try:
            vals = list(coords)
        except TypeError:
            vals = [float(coords)]
    flat: list[float] = []
    for v in vals:
        try:
            flat.append(float(v))
        except (TypeError, ValueError):
            continue
    if len(flat) < 3:
        return []
    # Pad if length not divisible by 3 (rare COM quirk)
    while len(flat) % 3 != 0:
        flat.append(0.0)
    return [(flat[i], flat[i + 1], flat[i + 2]) for i in range(0, len(flat), 3)]


def _entity_vertices(ent) -> list[tuple[float, float, float]]:
    def _read():
        return _coords_to_vertices(ent.Coordinates)

    try:
        verts = _com_retry(_read, label="Coordinates")
        if verts:
            return verts
    except Exception:
        pass

    # Fallback: Coordinate property / indexer if available
    try:

        def _by_index():
            out: list[tuple[float, float, float]] = []
            n = int(ent.Coordinates.Count) if hasattr(ent.Coordinates, "Count") else 0
            # Some builds expose Coordinates only as a flat SAFEARRAY (handled above)
            if n <= 0:
                return out
            for i in range(n):
                pt = ent.Coordinate(i) if hasattr(ent, "Coordinate") else ent.Coordinates.Item(i)
                out.append((float(pt[0]), float(pt[1]), float(pt[2])))
            return out

        return _com_retry(_by_index, label="Coordinate")
    except Exception:
        return []


def find_3d_polylines(doc) -> list[tuple[str, object, list[tuple[float, float, float]]]]:
    """Return list of (handle, entity, vertices) for AcDb3dPolyline in model space."""

    def _scan():
        ms = doc.ModelSpace
        count = int(ms.Count)
        found: list[tuple[str, object, list[tuple[float, float, float]]]] = []
        for i in range(count):
            ent = ms.Item(i)
            try:
                obj_name = str(ent.ObjectName)
            except Exception:
                continue
            if obj_name != OBJ_3DPOLY:
                continue
            try:
                handle = str(ent.Handle)
            except Exception:
                handle = f"IDX{i}"
            verts = _entity_vertices(ent)
            found.append((handle, ent, verts))
        return found

    return _com_retry(_scan, label="ModelSpace.scan", retries=40, delay=0.25)


def write_geometry_csv(
    out_folder: Path,
    handle: str,
    vertices: list[tuple[float, float, float]],
) -> Path:
    safe_id = "".join(c if c.isalnum() or c in "-_" else "_" for c in handle)
    out_path = out_folder / f"{safe_id}_Geometry.csv"
    with out_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["vertex", "x", "y", "z"])
        for idx, (x, y, z) in enumerate(vertices, start=1):
            # HK Grid metres: 4 dp = 0.1 mm (keeps accuracy, trims CAD float noise)
            writer.writerow(
                [idx, f"{round(x, 4):.4f}", f"{round(y, 4):.4f}", f"{round(z, 4):.4f}"]
            )
    return out_path


def main() -> int:
    print("AutoCAD 3D polyline geometry export")
    print("Uses AutoCAD COM. Packages must be installed in project .venv only.\n")

    dwg = prompt_dwg_path()
    out_folder = prompt_output_folder(dwg.parent)

    print("\nStarting AutoCAD (if not already running)...")
    acad = get_acad()
    print(f"Opening {dwg.name}...")
    doc = open_dwg(acad, dwg, read_only=True)

    try:
        print("Scanning model space for AcDb3dPolyline...")
        polylines = find_3d_polylines(doc)
        if not polylines:
            print("No 3D polylines (AcDb3dPolyline) found in model space.")
            return 1

        print(f"Found {len(polylines)} 3D polyline(s). Exporting one CSV per Handle...\n")
        print(f"{'Handle':<16} {'Vertices':>8}  Output")
        print("-" * 72)

        for handle, _ent, verts in polylines:
            if not verts:
                print(f"{handle:<16} {0:>8}  (skipped — no coordinates)")
                continue
            out_path = write_geometry_csv(out_folder, handle, verts)
            print(f"{handle:<16} {len(verts):>8}  {out_path}")

        print(f"\nDone. CSV files in: {out_folder}")
        return 0
    finally:
        _close_doc(doc)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nCancelled.")
        raise SystemExit(130)
