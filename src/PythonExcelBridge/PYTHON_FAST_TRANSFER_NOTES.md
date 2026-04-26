# PythonExcelBridge fast transfer notes

This build ports the RExcelBridge fast-transfer architecture to Python.

## New / updated public functions

- `PyObj(name)` marks an existing Python object so `PCall()` passes the object itself instead of a literal string.
- `PGetNumeric(name)` returns a NumPy numeric vector/array using the fast binary double path.
- `PSet(name, range)` now uses the fast binary double path automatically for fully numeric Excel ranges.
- `PSetTable(name, range, hasHeaders)` sends an Excel range to Python as a pandas DataFrame using typed columns.
- `PGetTable(name)` returns a pandas DataFrame using the typed table path.
- `PCall()` auto-dispatches NumPy numeric arrays and pandas DataFrames through fast return paths.
- `PLastTransfer()` reports the most recent transfer path and dimensions.

## Quick smoke tests

```excel
=PPing()
```

```excel
=PEval("import numpy as np")
=PEval("x = np.random.randn(10000,20)")
=PGetNumeric("x")
=PLastTransfer()
```

```excel
=PEval("def make_matrix(n,p):\n    import numpy as np\n    return np.random.randn(int(n), int(p))")
=PCall("make_matrix", 10000, 20)
=PLastTransfer()
```

For Excel -> Python numeric transfer, put this in a sheet cell, for example D10:

```excel
=RANDARRAY(10000,20)
```

Then:

```excel
=PSet("y", D10:W10009)
=PCall("numpy.shape", PyObj("y"))
=PLastTransfer()
```

For table/DataFrame transfer:

```excel
=PSetTable("df", D10:F13, TRUE)
=PCall("list", PyObj("df.columns"))
=PGetTable("df")
=PLastTransfer()
```

## Notes

- Fast numeric transfer requires NumPy.
- Fast table transfer requires pandas.
- Python uses row-major order, so the numeric path is more direct than R's matrix path.
- `"df"` is still a literal string. Use `PyObj("df")` when you want to pass an existing Python object to `PCall()`.
