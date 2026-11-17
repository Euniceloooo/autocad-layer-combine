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
BUILD_ID = "com-retry-4"
# AutoCAD busy / not ready for COM (common while opening/closing DWGs)
_RPC_E_CALL_REJECTED = -2147418111
_RPC_E_SERVERCALL_RETRYLATER = -2147417846
_COM_BUSY = {_RPC_E_CALL_REJECTED, _RPC_E_SERVERCALL_RETRYLATER}

# #region agent log
_DEBUG_LOG_PATH = Path(__file__).resolve().parent / ".cursor" / "debug-8178ab.log"


def _agent_log(hypothesis_id: str, location: str, message: str, data: dict | None = None) -> None:
    payload = {
        "sessionId": "8178ab",
        "runId": "com-retry-4",
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


def _com_hresult(ex: BaseException) -> int | None:
    try:
        return int(ex.args[0])
    except Exception:
        return None


def _com_retry(fn, *, retries: int = 50, delay: float = 0.2, label: str = "com"):
    """Retry when AutoCAD returns CALL_REJECTED / SERVERCALL_RETRYLATER."""
    last: BaseException | None = None
    for attempt in range(1, retries + 1):
        try:
            return fn()
        except Exception as ex:
            last = ex
            if _com_hresult(ex) not in _COM_BUSY:
                raise
            # #region agent log
            if attempt == 1 or attempt % 10 == 0:
                _agent_log(
                    "H4",
                    "main.py:_com_retry",
                    "COM busy, retrying",
                    {"label": label, "attempt": attempt, "hresult": _com_hresult(ex)},
                )
                print(f"  (AutoCAD busy — retry {attempt}/{retries}: {label})")
            # #endregion
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


class _AcadComMessageFilter:
    """OLE message filter: auto-retry when AutoCAD rejects a call as busy."""

    def HandleInComingCall(self, dwCallType, hTaskCaller, dwTickCount, lpInterfaceInfo):
        return 1  # SERVERCALL_ISHANDLED

    def RetryRejectedCall(self, hTaskCallee, dwTickCount, dwRejectType):
        # Retry for up to ~30s; return milliseconds until next retry
        if dwTickCount < 30000:
            return 250
        return -1  # give up

    def MessagePending(self, hTaskCallee, dwTickCount, dwPendingType):
        return 2  # PENDINGMSG_WAITDEFPROCESS


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
    # Wait until AutoCAD accepts a simple COM call
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
    """
    Open a DWG and return a real Document.

    Late-bound Documents.Open often does not return a Document (win32com leaves
    the Open method object). Prefer ActiveDocument after Open. Retries when
    AutoCAD is busy (RPC_E_CALL_REJECTED).
    """
    path_str = str(Path(path).resolve())
    expected = Path(path_str).name.lower()

    # Reuse if already open (avoids reject while re-opening the same file)
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
        # #region agent log
        _agent_log(
            "H5",
            "main.py:open_dwg",
            "reusing already-open document",
            {"path": path_str, "name": str(existing.Name)},
        )
        # #endregion
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
        opened_ok = _is_document(opened)
    except Exception:
        opened_ok = False
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

    # ActiveDocument may reject while AutoCAD finishes loading — retry + name match
    doc = None
    for attempt in range(1, 51):
        try:
            doc = acad.ActiveDocument
            if _is_document(doc):
                name = str(doc.Name).lower()
                full = str(getattr(doc, "FullName", "")).lower().replace("/", "\\")
                if name == expected or full == path_str.lower():
                    break
                # Wrong active doc — try find by name again
                found = _find_open()
                if found is not None and _is_document(found):
                    doc = found
                    break
        except Exception as ex:
            if _com_hresult(ex) not in _COM_BUSY:
                raise
            # #region agent log
            if attempt == 1 or attempt % 10 == 0:
                _agent_log(
                    "H4",
                    "main.py:open_dwg",
                    "ActiveDocument busy",
                    {"path": path_str, "attempt": attempt, "hresult": _com_hresult(ex)},
                )
            # #endregion
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
            "activeUsername": getattr(doc, "_username_", None) if doc is not None else None,
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


def _list_layer_names(doc) -> list[str]:
    """Read layer names with COM-busy retries (Layers.Item often rejects mid-scan)."""

    def _do() -> list[str]:
        layers = doc.Layers
        count = int(layers.Count)
        out: list[str] = []
        for i in range(count):
            layer = layers.Item(i)
            name = str(layer.Name)
            if name.upper() in ("0", "DEFPOINTS"):
                continue
            out.append(name)
        return out

    return _com_retry(_do, label="Layers.Item", retries=60, delay=0.25)


def collect_layer_names(acad, dwg_paths: list[Path]) -> list[str]:
    """Union of layer names for the picker.

    Only the first two DWGs are opened. Runtime logs showed the layer union
    stabilizes by file 2 (13 → 15), while opening all 10 triggers CALL_REJECTED
    mid-scan. Copy still opens every file and matches kept names there.
    """
    names: set[str] = set()
    scan_paths = dwg_paths[:2]
    print(
        f"Scanning layers from {len(scan_paths)} of {len(dwg_paths)} drawing(s) "
        "(enough for the layer list)..."
    )
    for idx, path in enumerate(scan_paths, 1):
        print(f"  [{idx}/{len(scan_paths)}] {path.name}")
        try:
            doc = open_dwg(acad, path, read_only=True)
        except Exception as ex:
            # #region agent log
            _agent_log(
                "H7",
                "main.py:collect_layer_names",
                "open failed during layer scan",
                {"path": str(path.name), "error": repr(ex)},
            )
            # #endregion
            print(f"  Warning: open failed for {path.name}: {ex}")
            continue
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
            try:
                for name in _list_layer_names(doc):
                    names.add(name)
            except Exception as ex:
                # #region agent log
                _agent_log(
                    "H6",
                    "main.py:collect_layer_names",
                    "layer scan failed, skipping file",
                    {"path": str(path.name), "error": repr(ex)},
                )
                # #endregion
                print(f"  Warning: could not read layers from {path.name}: {ex}")
            else:
                # #region agent log
                _agent_log(
                    "H3",
                    "main.py:collect_layer_names",
                    "layer scan ok",
                    {"path": str(path.name), "unionCount": len(names)},
                )
                # #endregion
        finally:
            _close_doc(doc)
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


_ORIGIN_LOCAL = Path(__file__).resolve().parent / "origin.local.json"


def load_origin_local() -> tuple[float, float, float] | None:
    """Load E/N/Z from gitignored origin.local.json if present."""
    if not _ORIGIN_LOCAL.is_file():
        return None
    try:
        data = json.loads(_ORIGIN_LOCAL.read_text(encoding="utf-8"))
        e = float(data["easting"])
        n = float(data["northing"])
        z = float(data.get("elevation", 0.0))
        return e, n, z
    except Exception as ex:
        print(f"Warning: could not read {_ORIGIN_LOCAL.name}: {ex}")
        return None


def prompt_origin() -> tuple[float, float]:
    """Return E/N for the origin mark. Prefers origin.local.json when present."""
    loaded = load_origin_local()
    if loaded is not None:
        e, n, z = loaded
        print(
            f"\nUsing origin from {_ORIGIN_LOCAL.name}: "
            f"E={e}  N={n}  Z={z} (cross uses Z=0)"
        )
        return e, n

    print(
        "\nOrigin mark (1 m cross). Enter Easting (E) and Northing (N) in METRES. "
        "Z will be 0. CAD units should already be metres."
    )
    print(
        f"(Tip: copy origin.example.json → {_ORIGIN_LOCAL.name} to skip this prompt; "
        "that file is gitignored.)"
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
    """Copy kept-layer entities into target with names `{layer}_{suffix}`.

    Renames layers in the open source document in memory, then CopyObjects so
    the target receives correctly named layers. Closes without saving so the
    DWG on disk is unchanged.
    """
    kept_lower = {k.lower() for k in kept_layers}
    # Writable so layer.Name can be changed; Close(False) discards disk writes.
    source = open_dwg(acad, source_path, read_only=False)
    copied = 0
    try:
        # Snapshot layer names first — renaming mutates the collection.
        src_names = _list_layer_names(source)
        renamed: dict[str, str] = {}  # original -> suffixed
        for src_name in src_names:
            if src_name.lower() not in kept_lower:
                continue
            expected = f"_{suffix}"
            if src_name.endswith(expected) or src_name.lower().endswith(expected.lower()):
                dst_name = src_name
            else:
                dst_name = f"{src_name}_{suffix}"
            if dst_name != src_name:
                try:
                    # Must not be the current layer when renaming in some AutoCAD builds
                    try:
                        if str(source.ActiveLayer.Name) == src_name:
                            source.ActiveLayer = source.Layers.Item("0")
                    except Exception:
                        pass
                    source.Layers.Item(src_name).Name = dst_name
                except Exception as ex:
                    print(f"  Layer rename failed {src_name!r} -> {dst_name!r}: {ex}")
                    continue
            renamed[src_name] = dst_name
            ensure_layer(target_doc, dst_name)

        # #region agent log
        _agent_log(
            "L1",
            "main.py:copy_kept_entities",
            "source layer rename map",
            {"file": source_path.name, "suffix": suffix, "renamed": renamed},
        )
        # #endregion

        if not renamed:
            return 0

        dst_names = set(renamed.values())
        ms = source.ModelSpace
        to_copy = []
        for i in range(ms.Count):
            ent = ms.Item(i)
            try:
                layer_name = str(ent.Layer)
            except Exception:
                continue
            if layer_name in dst_names:
                to_copy.append(ent)

        if not to_copy:
            # #region agent log
            _agent_log(
                "L2",
                "main.py:copy_kept_entities",
                "no entities on renamed layers",
                {"file": source_path.name, "dstNames": sorted(dst_names)},
            )
            # #endregion
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

        # Entities already carry suffixed layer names from the renamed source.
        # Soft-verify / correct any that somehow kept an old name.
        reverse = {v: v for v in renamed.values()}
        for new_ent in new_objs:
            try:
                ln = str(new_ent.Layer)
                if ln in reverse:
                    copied += 1
                    continue
                # Fallback: map original name if CopyObjects restored it
                if ln in renamed:
                    new_ent.Layer = renamed[ln]
                copied += 1
            except Exception:
                copied += 1

        if not new_objs and to_copy:
            copied = len(to_copy)

        # #region agent log
        _agent_log(
            "L2",
            "main.py:copy_kept_entities",
            "copy complete",
            {
                "file": source_path.name,
                "suffix": suffix,
                "entityCount": copied,
                "layers": sorted(set(renamed.values())),
            },
        )
        # #endregion
    finally:
        _close_doc(source)

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
    print("Uses AutoCAD COM. Packages must be installed in project .venv only.")
    print(f"Build: {BUILD_ID} (layer picker scans 2 DWGs only)")
    print(f"Script: {Path(__file__).resolve()}\n")
    # #region agent log
    _agent_log(
        "H0",
        "main.py:main",
        f"run start {BUILD_ID}",
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
