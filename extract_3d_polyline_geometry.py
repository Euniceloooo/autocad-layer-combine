"""
Extract 3D polyline vertices from one AutoCAD DWG into per-layer CSV files.

Each AcDb3dPolyline (model space or inside block inserts) becomes
`{Layer}_Geometry.csv` with columns: vertex,x,y,z

Assumes one 3D polyline per layer; warns and skips if a layer has more than one.
Requires AutoCAD installed. Run from the project .venv only.
"""

from __future__ import annotations

import csv
import math
import sys
import time
from collections import defaultdict
from dataclasses import dataclass
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

OBJ_3DPOLY = "AcDb3dPolyline"
OBJ_BLOCK_REF = "AcDbBlockReference"
_MAX_BLOCK_DEPTH = 8


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


@dataclass
class FoundPoly:
    layer: str
    handle: str
    vertices: list[tuple[float, float, float]]
    source: str


def _safe_filename(name: str) -> str:
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in name) or "layer"


def _as_xyz(pt) -> tuple[float, float, float]:
    try:
        return float(pt[0]), float(pt[1]), float(pt[2])
    except Exception:
        return float(pt.X), float(pt.Y), float(pt.Z)


def _block_ref_name(bref) -> str:
    for attr in ("EffectiveName", "Name"):
        try:
            name = str(getattr(bref, attr))
            if name:
                return name
        except Exception:
            continue
    return ""


def _transform_from_block_ref(bref):
    """Return f(x,y,z)->WCS for a typical insert (scale + Z rotation + insertion)."""
    ins = _as_xyz(bref.InsertionPoint)
    try:
        sx = float(bref.XScaleFactor)
        sy = float(bref.YScaleFactor)
        sz = float(bref.ZScaleFactor)
    except Exception:
        sx = sy = sz = 1.0
    try:
        rot = float(bref.Rotation)  # radians
    except Exception:
        rot = 0.0
    cos_r = math.cos(rot)
    sin_r = math.sin(rot)

    # If Normal is not world Z, build a simple orthonormal frame (OCS → WCS)
    try:
        nx, ny, nz = _as_xyz(bref.Normal)
    except Exception:
        nx, ny, nz = 0.0, 0.0, 1.0
    nlen = math.sqrt(nx * nx + ny * ny + nz * nz) or 1.0
    nx, ny, nz = nx / nlen, ny / nlen, nz / nlen
    world_z = abs(nx) < 1e-9 and abs(ny) < 1e-9 and abs(nz - 1.0) < 1e-6

    if world_z:

        def _map(x: float, y: float, z: float) -> tuple[float, float, float]:
            x1, y1, z1 = x * sx, y * sy, z * sz
            xr = x1 * cos_r - y1 * sin_r
            yr = x1 * sin_r + y1 * cos_r
            return (xr + ins[0], yr + ins[1], z1 + ins[2])

        return _map

    # Arbitrary normal: Azimuth/Elevation style axes (AutoCAD OCS)
    # Arbitrary axis algorithm (simplified)
    if abs(nx) < 1.0 / 64.0 and abs(ny) < 1.0 / 64.0:
        ax, ay, az = nz, 0.0, -nx
    else:
        ax, ay, az = -ny, nx, 0.0
    alen = math.sqrt(ax * ax + ay * ay + az * az) or 1.0
    ax, ay, az = ax / alen, ay / alen, az / alen
    bx = ny * az - nz * ay
    by = nz * ax - nx * az
    bz = nx * ay - ny * ax

    def _map(x: float, y: float, z: float) -> tuple[float, float, float]:
        x1, y1, z1 = x * sx, y * sy, z * sz
        xr = x1 * cos_r - y1 * sin_r
        yr = x1 * sin_r + y1 * cos_r
        zr = z1
        X = ax * xr + bx * yr + nx * zr + ins[0]
        Y = ay * xr + by * yr + ny * zr + ins[1]
        Z = az * xr + bz * yr + nz * zr + ins[2]
        return (X, Y, Z)

    return _map


def _compose_transforms(outer, inner):
    def _map(x: float, y: float, z: float) -> tuple[float, float, float]:
        return outer(*inner(x, y, z))

    return _map


def _identity_transform():
    def _map(x: float, y: float, z: float) -> tuple[float, float, float]:
        return (x, y, z)

    return _map


def _apply_transform(
    verts: list[tuple[float, float, float]],
    xform,
) -> list[tuple[float, float, float]]:
    return [xform(x, y, z) for x, y, z in verts]


def _scan_container(
    doc,
    container,
    *,
    xform,
    source: str,
    depth: int,
    visiting: set[str],
    found: list[FoundPoly],
) -> None:
    """Scan a ModelSpace / Block for 3D polys and nested block refs."""
    try:
        count = int(container.Count)
    except Exception:
        return
    for i in range(count):
        try:
            ent = container.Item(i)
            obj_name = str(ent.ObjectName)
        except Exception:
            continue

        if obj_name == OBJ_3DPOLY:
            try:
                layer = str(ent.Layer)
            except Exception:
                layer = "0"
            try:
                handle = str(ent.Handle)
            except Exception:
                handle = f"IDX{i}"
            local = _entity_vertices(ent)
            verts = _apply_transform(local, xform) if local else []
            found.append(
                FoundPoly(layer=layer, handle=handle, vertices=verts, source=source)
            )
            continue

        if obj_name != OBJ_BLOCK_REF or depth >= _MAX_BLOCK_DEPTH:
            continue

        bname = _block_ref_name(ent)
        # Skip unnamed / anonymous blocks (*U, *D, *Model_Space, etc.)
        if not bname or bname.startswith("*"):
            continue
        if bname in visiting:
            continue

        try:
            block_def = doc.Blocks.Item(bname)
        except Exception:
            continue

        child_xform = _compose_transforms(xform, _transform_from_block_ref(ent))
        visiting.add(bname)
        _scan_container(
            doc,
            block_def,
            xform=child_xform,
            source=f"block:{bname}" if source == "model" else f"{source}/{bname}",
            depth=depth + 1,
            visiting=visiting,
            found=found,
        )
        visiting.discard(bname)


def find_3d_polylines(doc) -> list[FoundPoly]:
    """Find AcDb3dPolyline in model space and inside block inserts (WCS)."""

    def _scan() -> list[FoundPoly]:
        found: list[FoundPoly] = []
        _scan_container(
            doc,
            doc.ModelSpace,
            xform=_identity_transform(),
            source="model",
            depth=0,
            visiting=set(),
            found=found,
        )
        return found

    return _com_retry(_scan, label="ModelSpace+blocks.scan", retries=40, delay=0.25)


def write_geometry_csv(
    out_folder: Path,
    layer_name: str,
    vertices: list[tuple[float, float, float]],
) -> Path:
    safe = _safe_filename(layer_name)
    out_path = out_folder / f"{safe}_Geometry.csv"
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
        print("Scanning model space + block inserts for AcDb3dPolyline...")
        polylines = find_3d_polylines(doc)
        if not polylines:
            print("No 3D polylines (AcDb3dPolyline) found.")
            return 1

        by_layer: dict[str, list[FoundPoly]] = defaultdict(list)
        for poly in polylines:
            by_layer[poly.layer].append(poly)

        print(
            f"Found {len(polylines)} 3D polyline(s) across {len(by_layer)} layer(s).\n"
        )
        print(f"{'Layer':<28} {'Polys':>5} {'Verts':>8}  Status")
        print("-" * 78)

        exported = 0
        skipped = 0
        for layer in sorted(by_layer.keys(), key=str.lower):
            items = by_layer[layer]
            if len(items) > 1:
                handles = ", ".join(p.handle for p in items)
                sources = ", ".join(sorted({p.source for p in items}))
                print(
                    f"{layer:<28} {len(items):>5} {'-':>8}  "
                    f"WARNING: >1 3D polyline — skipped "
                    f"(handles={handles}; sources={sources})"
                )
                skipped += 1
                continue

            poly = items[0]
            if not poly.vertices:
                print(
                    f"{layer:<28} {1:>5} {0:>8}  skipped — no coordinates "
                    f"({poly.source}, handle={poly.handle})"
                )
                skipped += 1
                continue

            out_path = write_geometry_csv(out_folder, layer, poly.vertices)
            print(
                f"{layer:<28} {1:>5} {len(poly.vertices):>8}  {out_path.name} "
                f"[{poly.source}]"
            )
            exported += 1

        print(f"\nExported {exported} CSV(s); skipped {skipped} layer(s).")
        print(f"Output folder: {out_folder}")
        return 0 if exported > 0 else 1
    finally:
        _close_doc(doc)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nCancelled.")
        raise SystemExit(130)
