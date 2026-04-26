# PythonExcelBridge worker
# JSON protocol over stdin/stdout, one request per line, one response per line.

from __future__ import annotations

import importlib
import io
import json
import math
import os
import runpy
import sys
import traceback
from pathlib import Path
from typing import Any

OBJECT_STORE: dict[str, Any] = {}
PLOTS_LOADED = False
BASE_DIR = Path(__file__).resolve().parent


def ensure_pandas_alias() -> None:
    try:
        import pandas as pd  # type: ignore
        globals()["pd"] = pd
    except Exception:
        pass



def make_ok(request_id: Any, result: Any) -> dict[str, Any]:
    return {"id": request_id, "ok": True, "result": result}


def make_err(request_id: Any, msg: Any) -> dict[str, Any]:
    return {"id": request_id, "ok": False, "error": str(msg)}


def is_scalar_json_value(x: Any) -> bool:
    return x is None or isinstance(x, (int, float, str, bool))


def is_json_matrix_like(x: Any) -> bool:
    if not isinstance(x, list) or not x:
        return False
    if not all(isinstance(row, list) and row for row in x):
        return False
    ncols = len(x[0])
    if not all(len(row) == ncols for row in x):
        return False
    return all(all(is_scalar_json_value(v) for v in row) for row in x)


def common_scalar_type(values: list[Any]) -> str:
    if all(v is None or isinstance(v, bool) for v in values):
        return "bool"
    if all(v is None or (isinstance(v, (int, float)) and not isinstance(v, bool)) for v in values):
        return "number"
    if all(v is None or isinstance(v, str) for v in values):
        return "string"
    return "any"


def vector_from_json(values: list[Any]) -> list[Any]:
    t = common_scalar_type(values)
    if t == "bool":
        return [False if v is None else bool(v) for v in values]
    if t == "number":
        return [math.nan if v is None else float(v) for v in values]
    if t == "string":
        return ["" if v is None else str(v) for v in values]
    return list(values)


def matrix_from_json(x: list[list[Any]]) -> list[list[Any]]:
    flat = [item for row in x for item in row]
    t = common_scalar_type(flat)
    if t == "bool":
        return [[False if v is None else bool(v) for v in row] for row in x]
    if t == "number":
        return [[math.nan if v is None else float(v) for v in row] for row in x]
    if t == "string":
        return [["" if v is None else str(v) for v in row] for row in x]
    return [list(row) for row in x]


def convert_json_value(x: Any) -> Any:
    if x is None:
        return None
    if is_scalar_json_value(x):
        return x
    if is_json_matrix_like(x):
        return matrix_from_json(x)
    if isinstance(x, list):
        converted = [convert_json_value(v) for v in x]
        if all(v is None or is_scalar_json_value(v) for v in converted):
            return vector_from_json(converted)
        return converted
    if isinstance(x, dict):
        if x.get("__pyexcel_arg_type") == "pyobj":
            nm = str(x.get("name", "")).strip()
            if not nm:
                raise ValueError("PyObj argument had a blank object name.")
            return resolve_object(nm)
        return {str(k): convert_json_value(v) for k, v in x.items()}
    return x


def normalize_call_args(args: Any) -> list[Any]:
    if isinstance(args, list):
        return [convert_json_value(v) for v in args]
    return [convert_json_value(args)]


def resolve_object(name: str) -> Any:
    if "." in name:
        obj: Any = globals()[name.split(".")[0]]
        for part in name.split(".")[1:]:
            obj = getattr(obj, part)
        return obj
    if name in globals():
        return globals()[name]
    if name in OBJECT_STORE:
        return OBJECT_STORE[name]
    raise KeyError(f"Object not found: {name}")


def resolve_function(fun_name: str):
    obj = resolve_object(fun_name)
    if not callable(obj):
        raise TypeError(f"Function not callable: {fun_name}")
    return obj


def coerce_for_json(x: Any) -> Any:
    if x is None or isinstance(x, (bool, int, float, str)):
        return x
    if isinstance(x, dict):
        return {str(k): coerce_for_json(v) for k, v in x.items()}
    if isinstance(x, tuple):
        return [coerce_for_json(v) for v in x]
    if isinstance(x, list):
        return [coerce_for_json(v) for v in x]
    try:
        import numpy as np  # type: ignore
        if isinstance(x, np.ndarray):
            return coerce_for_json(x.tolist())
    except Exception:
        pass
    try:
        import pandas as pd  # type: ignore
        if isinstance(x, pd.DataFrame):
            rows = [list(map(coerce_for_json, x.columns.tolist()))]
            rows.extend([[coerce_for_json(v) for v in row] for row in x.values.tolist()])
            return rows
        if isinstance(x, pd.Series):
            return [coerce_for_json(v) for v in x.tolist()]
    except Exception:
        pass
    if hasattr(x, "tolist"):
        try:
            return coerce_for_json(x.tolist())
        except Exception:
            pass
    if isinstance(x, io.IOBase):
        return str(x)
    return str(x)


def safe_length(x: Any) -> int:
    try:
        return len(x)
    except Exception:
        return 1


def format_dim(x: Any) -> str:
    try:
        import numpy as np  # type: ignore
        if isinstance(x, np.ndarray):
            return " x ".join(str(v) for v in x.shape)
    except Exception:
        pass
    if isinstance(x, list):
        if x and isinstance(x[0], list):
            return f"{len(x)} x {len(x[0])}"
        return f"{len(x)} x 1"
    return "1 x 1"


def object_summary_row(name: str, x: Any) -> list[str]:
    t = type(x)
    return [name, t.__name__, f"{t.__module__}.{t.__name__}", str(safe_length(x)), format_dim(x)]


def object_describe_table(name: str, x: Any) -> list[list[str]]:
    t = type(x)
    return [
        ["Field", "Value"],
        ["Name", name],
        ["Class", t.__name__],
        ["Type", f"{t.__module__}.{t.__name__}"],
        ["Length", str(safe_length(x))],
        ["Dimensions", format_dim(x)],
    ]


def ensure_matplotlib_loaded() -> None:
    global PLOTS_LOADED
    if PLOTS_LOADED:
        return
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    globals()["matplotlib"] = matplotlib
    globals()["plt"] = plt
    PLOTS_LOADED = True


def eval_code_for_excel(code: str) -> Any:
    local_env = globals()
    try:
        value = eval(code, local_env, local_env)
        return coerce_for_json(value)
    except SyntaxError:
        exec(code, local_env, local_env)
        return True


def render_plot_to_file(code: str, file: str, width: int = 800, height: int = 600, res: int = 96) -> str:
    if not file:
        raise ValueError("Plot file path is blank.")
    ensure_matplotlib_loaded()
    ensure_pandas_alias()
    Path(file).parent.mkdir(parents=True, exist_ok=True)
    dpi = max(int(res), 72)
    figsize = (max(int(width),1) / dpi, max(int(height),1) / dpi)
    plt.close('all')
    fig = None
    local_env = globals()
    local_env["FIGSIZE"] = figsize
    local_env["DPI"] = dpi
    local_env["PLOT_FILE"] = file
    try:
        candidate = eval(code, local_env, local_env)
        if hasattr(candidate, "savefig"):
            fig = candidate
    except SyntaxError:
        exec(code, local_env, local_env)
    if fig is None:
        fig = getattr(local_env.get("fig", None), "savefig", None) and local_env.get("fig") or None
    if fig is None:
        fig = plt.gcf()
    try:
        fig.set_size_inches(figsize[0], figsize[1], forward=True)
    except Exception:
        pass
    fig.savefig(file, dpi=dpi, bbox_inches="tight")
    plt.close(fig)
    return str(Path(file).resolve())


def source_file(path: str) -> bool:
    full = Path(path).expanduser()
    if not full.is_absolute():
        full = (BASE_DIR / path).resolve()
    code = full.read_text(encoding="utf-8")
    exec(compile(code, str(full), "exec"), globals(), globals())
    return True


# --- Fast transfer helpers added by ChatGPT ---
import time
import tempfile

LAST_TRANSFER: dict[str, Any] = {
    "method": "none",
    "source": "startup",
    "name": "",
    "class": "",
    "rows": 0,
    "cols": 0,
    "elapsed_ms": 0.0,
}


def _pxb_transfer_dir() -> Path:
    d = Path(tempfile.gettempdir()) / "PythonExcelBridgeTransfer"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _set_last_transfer(method: str, source: str, name: str, obj: Any, rows: int, cols: int, elapsed_ms: float) -> None:
    LAST_TRANSFER.update({
        "method": method,
        "source": source,
        "name": name or "",
        "class": obj.__class__.__name__ if obj is not None else "None",
        "rows": int(rows),
        "cols": int(cols),
        "elapsed_ms": round(float(elapsed_ms), 3),
    })


def last_transfer_table() -> list[list[Any]]:
    return [["Field", "Value"], *[[k, v] for k, v in LAST_TRANSFER.items()]]


def _resolve_named_object(name: str) -> Any:
    if name in OBJECT_STORE:
        return OBJECT_STORE[name]
    if name in globals():
        return globals()[name]
    raise KeyError(f"Object '{name}' was not found.")


def _is_numpy_numeric_array(x: Any) -> bool:
    try:
        import numpy as np  # type: ignore
        return isinstance(x, np.ndarray) and np.issubdtype(x.dtype, np.number)
    except Exception:
        return False


def _as_numeric_array(x: Any):
    import numpy as np  # type: ignore
    if isinstance(x, np.ndarray):
        if not np.issubdtype(x.dtype, np.number):
            raise TypeError("Object is not a numeric NumPy array.")
        arr = x.astype("float64", copy=False)
    elif isinstance(x, (list, tuple)):
        arr = np.asarray(x, dtype="float64")
    elif isinstance(x, (int, float)) and not isinstance(x, bool):
        arr = np.asarray([[float(x)]], dtype="float64")
    else:
        raise TypeError("Object is not numeric.")

    if arr.ndim == 0:
        arr = arr.reshape(1, 1)
    elif arr.ndim == 1:
        arr = arr.reshape(arr.shape[0], 1)
    elif arr.ndim == 2:
        pass
    else:
        raise TypeError("Only numeric scalars, vectors, and 2D arrays are supported for fast numeric transfer.")
    return arr


def export_numeric_object(x: Any, source: str = "numeric", name: str = "") -> dict[str, Any]:
    t0 = time.perf_counter()
    arr = _as_numeric_array(x)
    rows, cols = int(arr.shape[0]), int(arr.shape[1])
    file = _pxb_transfer_dir() / f"python_numeric_{os.getpid()}_{time.time_ns()}.bin"
    # C-order matches Excel row-major reconstruction in C#.
    arr.astype("float64", copy=False).ravel(order="C").tofile(str(file))
    elapsed = (time.perf_counter() - t0) * 1000
    _set_last_transfer("numeric", source, name, x, rows, cols, elapsed)
    return {
        "__pyexcel_transfer_type": "numeric",
        "file": str(file).replace("\\", "/"),
        "rows": rows,
        "cols": cols,
        "storage_order": "row-major-double64",
        "na": "NaN",
    }


def _is_pandas_dataframe(x: Any) -> bool:
    try:
        import pandas as pd  # type: ignore
        return isinstance(x, pd.DataFrame)
    except Exception:
        return False


def export_table_object(x: Any, source: str = "table", name: str = "") -> dict[str, Any]:
    import numpy as np  # type: ignore
    import pandas as pd  # type: ignore
    t0 = time.perf_counter()
    if not isinstance(x, pd.DataFrame):
        raise TypeError("Object is not a pandas DataFrame.")
    rows, cols = int(x.shape[0]), int(x.shape[1])
    columns: list[dict[str, Any]] = []
    transfer_dir = _pxb_transfer_dir()
    for col_name in x.columns:
        s = x[col_name]
        name_str = str(col_name)
        if pd.api.types.is_numeric_dtype(s):
            vals = s.astype("float64").to_numpy(copy=False)
            file = transfer_dir / f"python_table_col_{os.getpid()}_{time.time_ns()}_{len(columns)}.bin"
            np.asarray(vals, dtype="float64").tofile(str(file))
            columns.append({"name": name_str, "type": "numeric", "file": str(file).replace("\\", "/"), "na": "NaN"})
        elif pd.api.types.is_bool_dtype(s):
            vals = [None if pd.isna(v) else bool(v) for v in s.tolist()]
            columns.append({"name": name_str, "type": "logical", "values": vals})
        else:
            vals = [None if pd.isna(v) else str(v) for v in s.tolist()]
            columns.append({"name": name_str, "type": "character", "values": vals})
    elapsed = (time.perf_counter() - t0) * 1000
    _set_last_transfer("typed table", source, name, x, rows, cols, elapsed)
    return {
        "__pyexcel_transfer_type": "table",
        "rows": rows,
        "cols": cols,
        "include_headers": True,
        "columns": columns,
    }


def coerce_for_transport(x: Any, source: str = "general", name: str = "") -> Any:
    try:
        if _is_numpy_numeric_array(x):
            return export_numeric_object(x, source=source, name=name)
    except Exception:
        raise
    try:
        if _is_pandas_dataframe(x):
            return export_table_object(x, source=source, name=name)
    except Exception:
        raise
    return coerce_for_json(x)


def set_numeric_from_file(name: str, file: str, rows: int, cols: int) -> bool:
    import numpy as np  # type: ignore
    t0 = time.perf_counter()
    arr = np.fromfile(file, dtype="float64", count=int(rows) * int(cols)).reshape((int(rows), int(cols)), order="C")
    OBJECT_STORE[name] = arr
    globals()[name] = arr
    elapsed = (time.perf_counter() - t0) * 1000
    _set_last_transfer("numeric to Python", "PSet", name, arr, rows, cols, elapsed)
    return True


def set_table_from_payload(name: str, rows: int, cols: int, columns: list[dict[str, Any]]) -> bool:
    import numpy as np  # type: ignore
    import pandas as pd  # type: ignore
    t0 = time.perf_counter()
    data: dict[str, Any] = {}
    for col in columns:
        col_name = str(col.get("name", f"V{len(data)+1}"))
        col_type = str(col.get("type", "character"))
        if col_type == "numeric":
            file = str(col.get("file", ""))
            data[col_name] = np.fromfile(file, dtype="float64", count=int(rows))
        elif col_type == "logical":
            data[col_name] = col.get("values", [])
        else:
            data[col_name] = col.get("values", [])
    df = pd.DataFrame(data)
    OBJECT_STORE[name] = df
    globals()[name] = df
    elapsed = (time.perf_counter() - t0) * 1000
    _set_last_transfer("typed table to Python", "PSetTable", name, df, rows, cols, elapsed)
    return True
# --- end fast transfer helpers ---


def handle_request(req: dict[str, Any]) -> dict[str, Any]:
    request_id = req.get("id")
    cmd = req.get("cmd", "")
    try:
        if cmd == "ping":
            return make_ok(request_id, f"OK | PythonExcelBridge | Python {sys.version.split()[0]}")
        if cmd == "source":
            file = req.get("file")
            if not isinstance(file, str):
                raise ValueError("file is missing.")
            return make_ok(request_id, source_file(file))
        if cmd == "eval":
            code = req.get("code")
            if not isinstance(code, str):
                raise ValueError("code is missing.")
            result = eval_code_for_excel(code)
            return make_ok(request_id, coerce_for_transport(result, source="PEval", name=""))
        if cmd == "plot":
            code = req.get("code")
            file = req.get("file")
            width = int(req.get("width", 800))
            height = int(req.get("height", 600))
            res = int(req.get("res", 96))
            if not isinstance(code, str):
                raise ValueError("code is missing.")
            if not isinstance(file, str):
                raise ValueError("file is missing.")
            return make_ok(request_id, render_plot_to_file(code, file, width, height, res))
        if cmd == "call":
            fun_name = req.get("fun")
            if not isinstance(fun_name, str):
                raise ValueError("fun is missing.")
            fun = resolve_function(fun_name)
            args = normalize_call_args(req.get("args", []))
            result = fun(*args)
            return make_ok(request_id, coerce_for_transport(result, source="PCall", name=fun_name))
        if cmd == "set":
            name = req.get("name")
            if not isinstance(name, str):
                raise ValueError("name is missing.")
            value = convert_json_value(req.get("value"))
            OBJECT_STORE[name] = value
            globals()[name] = value
            return make_ok(request_id, True)
        if cmd == "get":
            name = req.get("name")
            if not isinstance(name, str):
                raise ValueError("name is missing.")
            if name in OBJECT_STORE:
                return make_ok(request_id, coerce_for_transport(OBJECT_STORE[name], source="PGet", name=name))
            if name in globals():
                return make_ok(request_id, coerce_for_transport(globals()[name], source="PGet", name=name))
            raise KeyError(f"Object '{name}' was not found.")
        if cmd == "set_numeric":
            name = req.get("name")
            file = req.get("file")
            rows = int(req.get("rows", 0))
            cols = int(req.get("cols", 0))
            if not isinstance(name, str) or not isinstance(file, str):
                raise ValueError("name or file is missing.")
            return make_ok(request_id, set_numeric_from_file(name, file, rows, cols))
        if cmd == "get_numeric":
            name = req.get("name")
            if not isinstance(name, str):
                raise ValueError("name is missing.")
            return make_ok(request_id, export_numeric_object(_resolve_named_object(name), source="PGetNumeric", name=name))
        if cmd == "set_table":
            name = req.get("name")
            rows = int(req.get("rows", 0))
            cols = int(req.get("cols", 0))
            columns = req.get("columns", [])
            if not isinstance(name, str) or not isinstance(columns, list):
                raise ValueError("name or columns is missing.")
            return make_ok(request_id, set_table_from_payload(name, rows, cols, columns))
        if cmd == "get_table":
            name = req.get("name")
            if not isinstance(name, str):
                raise ValueError("name is missing.")
            return make_ok(request_id, export_table_object(_resolve_named_object(name), source="PGetTable", name=name))
        if cmd == "last_transfer":
            return make_ok(request_id, last_transfer_table())
        if cmd == "exists":
            name = req.get("name")
            if not isinstance(name, str):
                raise ValueError("name is missing.")
            return make_ok(request_id, name in OBJECT_STORE or name in globals())
        if cmd == "remove":
            name = req.get("name")
            if not isinstance(name, str):
                raise ValueError("name is missing.")
            OBJECT_STORE.pop(name, None)
            globals().pop(name, None)
            return make_ok(request_id, True)
        if cmd == "objects":
            rows: list[list[str]] = [["Name", "Class", "Type", "Length", "Dimensions"]]
            for nm in sorted(OBJECT_STORE):
                rows.append(object_summary_row(nm, OBJECT_STORE[nm]))
            return make_ok(request_id, rows)
        if cmd == "describe":
            name = req.get("name")
            if not isinstance(name, str):
                raise ValueError("name is missing.")
            if name in OBJECT_STORE:
                return make_ok(request_id, object_describe_table(name, OBJECT_STORE[name]))
            if name in globals():
                return make_ok(request_id, object_describe_table(name, globals()[name]))
            raise KeyError(f"Object '{name}' was not found.")
        return make_err(request_id, f"Unknown command: {cmd}")
    except Exception as exc:
        return make_err(request_id, f"{exc}")


def main() -> None:
    startup_file = sys.argv[1] if len(sys.argv) >= 2 else None
    try:
        if startup_file and Path(startup_file).is_file():
            print(f"Including startup file: {startup_file}", file=sys.stderr, flush=True)
            source_file(startup_file)
            print("Startup file loaded successfully.", file=sys.stderr, flush=True)
        else:
            print("No startup file provided.", file=sys.stderr, flush=True)
    except Exception as exc:
        print(f"FATAL startup error: {traceback.format_exc()}", file=sys.stderr, flush=True)
        print(json.dumps(make_err(None, f"Startup error: {exc}")), flush=True)
        return

    print("Worker entering request loop.", file=sys.stderr, flush=True)
    while True:
        try:
            line = sys.stdin.readline()
        except Exception:
            print(f"Readline failed: {traceback.format_exc()}", file=sys.stderr, flush=True)
            break
        if line == "":
            break
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except Exception as exc:
            print(json.dumps(make_err(None, f"Invalid JSON: {exc}")), flush=True)
            continue
        try:
            resp = handle_request(req)
        except Exception:
            resp = make_err(None, traceback.format_exc())
        print(json.dumps(resp), flush=True)


if __name__ == "__main__":
    main()
