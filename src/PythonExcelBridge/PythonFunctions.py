from __future__ import annotations

import math


def CholDecomp(mat):
    import numpy as np
    arr = np.array(mat, dtype=float)
    upper = np.linalg.cholesky(arr).T
    return upper.tolist()


def CholDecompLower(mat):
    import numpy as np
    arr = np.array(mat, dtype=float)
    lower = np.linalg.cholesky(arr)
    return lower.tolist()


def ShowObjectInfo(x):
    dims = getattr(x, "shape", None)
    if dims is None:
        if isinstance(x, list) and x and isinstance(x[0], list):
            dims = (len(x), len(x[0]))
        elif isinstance(x, list):
            dims = (len(x),)
        else:
            dims = None
    dim_text = "x".join(str(v) for v in dims) if dims else "NULL"
    try:
        length = len(x)
    except Exception:
        length = 1
    return f"type = {type(x)}; dim = {dim_text}; length = {length}"


def MatrixMultiply(a, b):
    import numpy as np
    return (np.array(a, dtype=float) @ np.array(b, dtype=float)).tolist()


def IdentityMatrix(n):
    import numpy as np
    return np.eye(int(n), dtype=float).tolist()


def RowSums(mat):
    import numpy as np
    return np.array(mat, dtype=float).sum(axis=1).tolist()


def ReloadFunctionsPython():
    return "PythonFunctions.py reloaded"


def hello_python_bridge():
    return "Hello from PythonExcelBridge"
