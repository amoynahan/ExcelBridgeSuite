
# RExcelBridge Usage Guide

This guide explains how to use RExcelBridge, the R-based add-in in ExcelBridgeSuite.

RExcelBridge is the reference implementation of the bridge pattern:
Excel → Add-in → R → Add-in → Excel

If you understand this workflow, the Julia and Python bridges will follow naturally.

---

## Quick Start

### Step 1 — Check connectivity

=RPing()

![RPing](docs/images/Rping.jpg)

Expected:
OK | R version ...

---


### Step 2 — Evaluate a simple expression

=REval("1+1")

![Simple](docs/images/SimpleExpression.jpg)

Expected:
2

---

### Step 3 — Return multiple values

=REval("c(10,20,30)")

![Vector](docs/images/ReturnVector.jpg)

---

## Validation

If the steps above work, your environment is correctly configured and you are ready to continue.

The following have been validated:

- The Excel add-in is loaded and active  
- R is accessible via `Rscript.exe`  
- The bridge is executing R code successfully  
- All required files are correctly located in the publish folder  

---

## Core Workflow

### Evaluate R code

=REval("sqrt(16)")

---

### Call an R function

=RCall("sum",1,2,3)

---

### Return structured data

=REval("matrix(c(1,2,3,4), nrow=2)")

![Matrix](docs/images/ReturnMatrix.jpg)

---

### Pass data from Excel to R

=RSet("x", A1:B3)

![RSet](docs/images/RSet.jpg)

=RGet("x")

![RGet](docs/images/RGet.jpg)

---

## Custom Functions

Place in RFunctions.R

Example:

add_ten <- function(x) {
  x + 10
}

=REval("add_ten(5)")

![Add](docs/images/Add10.jpg)

---

## Numerical Example — Cholesky

## Cholesky Decomposition (Advanced Example)

This demonstrates real numerical computation.

### R Wrapper

Add the wrapper to `RFunctions.R`:

chol_decomp <- function(x, tol = 1e-8) {  
  x <- as.matrix(x)  
  
  if (!is.numeric(x)) {  
    stop("Input must be numeric.")  
  }  
  
  if (nrow(x) != ncol(x)) {  
    stop("Input matrix must be square.")  
  }  
  
  if (max(abs(x - t(x))) > tol) {  
    stop("Input matrix must be symmetric.")  
  }  
  
  chol(x)  
}

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
=RCall("CholDecomp", A1:B2)
```

Expected result:

```
2        1
0   1.414214
```
![Description of image](docs/images/CholDecomp.jpg)

---

## Plotting

RExcelBridge supports two plotting workflows:

1. Simple plotting using the RExcelBridge ribbon
2. Advanced/dynamic plotting using Excel macros

---

## Simple Plotting: Insert Plot from the Add-in Ribbon

Use this approach when you want to create a plot once and insert it into the worksheet.

### Example

=RPlot("plot(1:5, c(0,1,4,9,16), type='b')", "BasicPlot", 800, 600)

The formula returns the path to the generated PNG file.

![RPlot](docs/images/RPlot.jpg)

### Insert the plot

1. Select the cell containing the returned plot path
2. Go to the RExcelBridge ribbon
3. Click Insert Plot From Selected Cell

![Add-in Menu](docs/images/AddinMenu.jpg)

![Simple Plot](docs/images/SimplePlot.jpg)

---

## Advanced Plotting: Dynamic Plots with Excel Macros

Use this approach when you want the plot to update when the worksheet data changes.

This workflow uses two parts:

- RPlotDataNamed creates the plot and returns the image file path
- PlotLink displays the image in the worksheet

This requires the DisplayImages VBA macro.

### Why use this approach

This pattern is useful when:

- the plot depends on worksheet data
- you want to press F9 and refresh the plot
- you want a reusable plotting workflow
- you are using more complex R code such as ggplot2

### Add the VBA module

1. Open the VBA editor with Alt + F11
2. In the Project pane, right-click the workbook
3. Select Insert → Module
4. Rename the module to DisplayImages
5. Paste the PlotLink macro code into that module

### Save as macro-enabled workbook

Save the workbook as:

Excel Macro-Enabled Workbook (*.xlsm)

If you save as .xlsx, Excel will remove the VBA code.

### Step 1 — Create the plot

In one cell, use RPlotDataNamed to generate the plot and return the PNG path.

![RPlotDataNamed](docs/images/RPlotDataNamed.jpg)

### Step 2 — Display the plot

In another cell, use PlotLink to display the generated image.

![PlotLink](docs/images/PlotLink.jpg)

---

## Performance (Large Data)

See:
[Performance Guide](PERFORMANCE.md)

---

## Troubleshooting

Ping fails → reload add-in  
R fails → check rscript-path.txt  
Plot fails → check plot-path.txt  

---

## Function Reference

RPing()
REval(code)
RCall(fun,...)
RSet(name,value)
RGet(name)
RPlot(...)
RSource(file)
RObjects()
RDescribe(name)
