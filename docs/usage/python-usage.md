# PythonExcelBridge Usage Guide

This guide explains how to use PythonExcelBridge, the Python-based add-in in ExcelBridgeSuite.

PythonExcelBridge follows the bridge pattern:
Excel → Add-in → Python → Add-in → Excel

If you understand this workflow, the R and Julia bridges will follow naturally.

---

## Quick Start

### Step 1 — Check connectivity

=PPing()

Expected:
OK | Python version ...

---

### Step 2 — Evaluate a simple expression

=PEval("1+1")

Expected:
2

---

### Step 3 — Return multiple values

=PEval("[10,20,30]")

---

## Validation

If the steps above work, your environment is correctly configured and you are ready to continue.

The following have been validated:

- The Excel add-in is loaded and active  
- Python is accessible via the configured executable  
- The bridge is executing Python code successfully  
- All required files are correctly located in the publish folder  

---

## Core Workflow

### Evaluate Python code

=PEval("16**0.5")

---

### Call a Python function

=PEval("sum([1,2,3])")

---

### Return structured data

=PEval("import numpy as np; np.array([[1,2],[3,4]])")

---

### Pass data from Excel to Python

=PSet("x", A1:B3)

=PGet("x")

---

## Fast Data Transfer

PythonExcelBridge includes optimized transfer paths for numeric data and tables.

### Numeric matrices

=PSet("x", A1:D1000)

=PGetNumeric("x")

=PLastTransfer()

---

### Tables / DataFrames

=PSetTable("df", A1:D1000, TRUE)

=PGetTable("df")

=PLastTransfer()

---

### Large data example

=PSet("x", A1:Z10000)

=PGetNumeric("x")

=PLastTransfer()

---

## Custom Functions

Place in PythonFunctions.py

Example:

def add_ten(x):
    return x + 10

=PEval("add_ten(5)")

---

## Numerical Example — Cholesky

## Cholesky Decomposition (Advanced Example)

This demonstrates real numerical computation.

### Python Wrapper

Add the wrapper to `PythonFunctions.py`:

import numpy as np

def chol_decomp(x, tol=1e-8):
    x = np.array(x, dtype=float)
    
    if x.ndim != 2:
        raise ValueError("Input must be a matrix.")
    
    if x.shape[0] != x.shape[1]:
        raise ValueError("Input matrix must be square.")
    
    if np.max(np.abs(x - x.T)) > tol:
        raise ValueError("Input matrix must be symmetric.")
    
    return np.linalg.cholesky(x)

---

### Excel Example

### Requirements
- Matrix must be square
- Matrix must be symmetric
- Matrix must be positive definite

Put this matrix in Excel:

```
4   2
2   3
```

Then run:

```excel
=PSet("x", A1:B2)
=PEval("chol_decomp(x)")
```

Expected result:

```
2        1
0   1.414214
```

---

## Plotting

PythonExcelBridge supports plotting using matplotlib.

---

### Simple Plotting

=PPlot("import matplotlib.pyplot as plt; plt.plot([1,2,3],[1,4,9])")

The formula returns the path to the generated PNG file.

---

### Plot Excel data

=PPlotDataNamed(A1:A10,"x",B1:B10,"y","import matplotlib.pyplot as plt; plt.plot(x,y)")

---

### Dynamic Plotting

Use this approach when you want the plot to update when worksheet data changes.

This workflow uses:

- PPlotDataNamed to create the plot and return the image path  
- PlotLink (VBA macro) to display the image  

---

## Performance (Large Data)

See:
[Performance Guide](PERFORMANCE.md)

---

## Troubleshooting

Ping fails → reload add-in  
Python fails → check python-path.txt  
Plot fails → check plot-path.txt  

---

## Function Reference

PPing()
PEval(code)
PCall(fun,...)
PSet(name,value)
PGet(name)
PSetTable(name,value,hasHeaders)
PGetTable(name)
PGetNumeric(name)
PPlot(...)
PPlotDataNamed(...)
PLastTransfer()
PSource(file)
PObjects()
PDescribe(name)
