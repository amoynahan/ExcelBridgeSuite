# PythonExcelBridge Usage Guide

This guide explains how to use PythonExcelBridge, the Python-based add-in in ExcelBridgeSuite.

PythonExcelBridge follows the bridge pattern:
Excel → Add-in → Python → Add-in → Excel

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

If the steps above work, your environment is correctly configured.

The following have been validated:

- The Excel add-in is loaded and active  
- Python is accessible via the configured executable  
- The bridge is executing Python code successfully  

---

## Core Workflow

### Evaluate Python code

=PEval("16**0.5")

---

### Pass data from Excel to Python

=PSet("x", A1:B3)

=PGet("x")

---

## Fast Data Transfer

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

## Python Examples

### Lambda Function

=PEval("(lambda x: x + 10)(5)")

Expected:
15

---

### List Comprehension

=PEval("[x*x for x in [1,2,3,4]]")

Expected:
1   4   9   16

---

### Call python function (create random matrix)

loads the NumPy library into the Python session and assigns it the alias np.

=PEval("import numpy as np")

Generate a 10000 × 26 random matrix:

=PCall("np.random.rand",10000,26)


### Cholesky Decomposition (NumPy)

### Requirements
- Matrix must be square
- Matrix should be symmetric
- Matrix must be positive definite

Put this matrix in Excel:

```
4   2
2   3
```

### Step 1 — load the NumPy librar

=PEval("import numpy as np")

### Step 2 — Run Cholesky

=PCall("np.linalg.cholesky",A1:B2)

Expected result:

```
2        0
1   1.414214
```

Note:
NumPy returns a **lower triangular matrix**, which differs from R's upper triangular result.

## Simple Python Plot

The example below creates a matplotlib plot, saves it to the plot cache folder, and returns the PNG path to Excel.

### Excel Formula

```excel
=PPlot("plt.plot([1,2,3,4,5], [1,4,9,16,25], marker='o')", "BasicPlot", 800, 600)
```

### Result

The formula returns a PNG file path similar to:

```text
D:\OneDrive - Avalere Health\Documents\PythonExcelBridge\PlotCache\Book1_Sheet1_A4_BasicPlot.png
```

Select the cell containing the PNG path and use either:

```text
Add-ins -> PythonExcelBridge -> Insert Plot From Selected Cell
```

or:

```text
Ctrl + Shift + P
```

### Example Output

![Simple Plot](../images/python/SimplePlot.jpg)

## Dynamic Python Plot from Excel Data

This example demonstrates a refreshable matplotlib plot driven directly from worksheet data.

The workflow is:

```text
Excel worksheet data
    ↓
PPlotDataNamed()
    ↓
matplotlib generates PNG
    ↓
PlotLink() displays image
    ↓
Worksheet recalculation refreshes the plot
```

### 1. Worksheet Data

Create worksheet data in columns A and B.

Column A contains X values.

Column B contains formulas that change dynamically.

Example:

```text
A1: X      B1: Y
A2: 1      B2: =A2+(RAND()-0.5)
A3: 2      B3: =A3+(RAND()-0.5)
A4: 3      B4: =A4+(RAND()-0.5)
...
A11: 10    B11: =A11+(RAND()-0.5)
```

Because the Y values use `RAND()`, the data changes every time Excel recalculates.

### 2. Create the Python Plot

In cell `D1`, enter:

```excel
=PPlotDataNamed(
  "plt.plot(X, Y, marker='o')",
  "BasicPlotData",
  800,
  600,
  "X", A2:A11,
  "Y", B2:B11
)
```

This formula reads the Excel ranges into Python, creates a matplotlib plot, saves the PNG file, and returns the PNG path to Excel.

The returned value will look similar to:

```text
D:\OneDrive - Avalere Health\Documents\PythonExcelBridge\PlotCache\Test_Sheet1_D1_BasicPlotData.png
```

### 3. Display the Plot in Excel

In cell `D2`, enter:

```excel
=PlotLink(D1, 600, 400)
```

`PlotLink()` links the PNG path to a displayed image on the worksheet.

Arguments:

```text
PlotLink(plotPath, widthPx, heightPx)
```

### 4. Import the VBA Plot Display Module

Import the VBA module:

```text
vba/DisplayImages.bas
```

This module provides the `PlotLink()` function and the image refresh logic.

It handles:

- detecting `=PlotLink(...)` formulas
- inserting plot images
- refreshing images after recalculation
- replacing outdated images automatically

### 5. Add the Worksheet Recalculation Event

In the worksheet VBA code page, add:

```vba
Private Sub Worksheet_Calculate()
    RefreshPlotLinksInSheet Me
End Sub
```

This event handler refreshes displayed plot images whenever the worksheet recalculates.

### 6. Refresh the Plot

Press:

```text
F9
```

Each recalculation changes the worksheet data, reruns the Python plotting code, regenerates the PNG file, and refreshes the displayed plot image.

### Example Output

![Dynamic Plot](../images/python/DynamicPlot.jpg)

## Troubleshooting

Ping fails → reload add-in  
Python fails → check python-path.txt  

---

## Function Reference
```
PPing()
PEval(code)
PSet(name,value)
PGet(name)
PSetTable(name,value,hasHeaders)
PGetTable(name)
PGetNumeric(name)
PLastTransfer()
```
