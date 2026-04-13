# ExcelBridgeSuite Usage Guide

This guide walks through how to use ExcelBridgeSuite, starting with basic validation and progressing to real computation and plotting.

Examples focus on RExcelBridge, but the same concepts apply to JuliaExcelBridge and PythonExcelBridge.

---

## Where to Put Your R Functions

Custom R functions that you want to call from Excel should be placed in `RFunctions.R`.

This file is intended for user-defined wrappers, helper functions, and reusable logic. In general:

- `startup.R` is for startup behavior  
- `worker.R` is for bridge execution logic  
- `RFunctions.R` is for your custom Excel-callable R functions  

## Before You Start

Make sure:

- The add-in is attached in Excel  
- The add-in is checked under Excel Add-ins  
- For RExcelBridge, rscript-path.txt points to a valid Rscript.exe  
- All files remain together in the publish folder  

---

## 1. Start with Ping

The first thing to test is Ping().

### Purpose

Ping() is a connectivity and environment check. It confirms that:

- Excel has loaded the add-in  
- The bridge is responding  
- R is accessible  
- You are running the expected version  

### Example

=Ping()

### Expected Result

You should see something like:

OK | R version 4.5.3 (2026-03-11 ucrt)

### What it means

- OK means the add-in is loaded and responding  
- The R version confirms that R is found and executable  
- The version string confirms your runtime environment  

This verifies the full pipeline:

Excel → Add-in → R → Add-in → Excel  

### If it fails

- Check the add-in is attached and checked  
- Restart Excel and reload the .xll  
- Verify rscript-path.txt  
- Confirm Rscript.exe runs from command line  

---

## 2. Evaluate a Simple Expression

Test that R can execute code.

Example:

=REval("1+1")

Expected result:

2

Another example:

=REval("sqrt(16)")

Expected result:

4

---

## 3. Return a Vector

Test returning multiple values.

Example:

=REval("c(10,20,30)")

Expected:

Values returned to Excel. These may spill across multiple cells depending on how the bridge returns arrays.

---

## 4. Return a Matrix

Test 2D data handling.

Example:

=REval("matrix(c(1,2,3,4), nrow=2)")

Expected:

A 2 × 2 result returned to Excel.

---

## 5. Pass Data from Excel to R

This allows Excel to act as a front end.

Example data in Excel (A1:B3):

1   2  
3   4  
5   6  

Example workflow:

=RPut("x", A1:B3)  
=REval("x")

---

## 6. Create a Simple R Wrapper

Define reusable logic in R.

User-defined functions should be added to `RFunctions.R`.

This file is intended to hold custom R functions that you want to call from Excel. Keeping user functions in `RFunctions.R` makes them easier to find, maintain, and extend.

Example in `RFunctions.R`:

add_ten <- function(x) {  
  x + 10  
}

Call from Excel:

=REval("add_ten(5)")

Expected result:

15

---

## 7. Cholesky Decomposition (Advanced Example)

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

Put this matrix in Excel:

4   2  
2   3  

Then run:

=RPut("x", A1:B2)  
=REval("chol_decomp(x)")

Expected result:

2        1  
0   1.414214  

### Requirements

- Matrix must be square  
- Matrix must be symmetric  
- Matrix must be positive definite  

---

## 8. Plotting

Test graphical output.

Example:

=RPlot("plot(1:10)")

or:

=RPlotTest()

### What happens

1. R creates a PNG file  
2. The file is saved to disk  
3. The bridge returns or inserts the plot into Excel  

---

## 9. plot-path.txt Behavior

Controls where plots are saved.

If a path is provided, for example:

C:\Users\YourName\Documents\RExcelBridge\Plots

- Plots are written to this folder  
- The folder will be created if needed  

If plot-path.txt is empty or missing:

- A default location is used  
- Typically a Documents folder or a temporary directory  

For consistent results, it is recommended to set this explicitly.

---

## 10. Example Plot Wrapper in R

User-defined plotting helpers can also be added to `RFunctions.R`.

Example:

make_plot <- function(file) {  
  png(file, width = 800, height = 600)  
  plot(1:10, 1:10, main = "Basic Plot")  
  dev.off()  
  file  
}

Call from Excel:

=REval("make_plot('test.png')")

---

## 11. Suggested Workflow

Test in this order:

1. =Ping()  
2. =REval("1+1")  
3. Return a vector  
4. Return a matrix  
5. Pass Excel range to R  
6. Create wrapper function  
7. Run Cholesky example  
8. Generate a plot  

---

## 12. Troubleshooting

Ping fails  
- Add-in not loaded  
- Wrong .xll  
- Excel needs restart  

R execution fails  
- Check rscript-path.txt  
- Verify R installation  

Matrix issues  
- Ensure 2D shape is preserved  
- Ensure numeric input  

Cholesky fails  
- Matrix not symmetric  
- Matrix not positive definite  

Plot fails  
- Check plot-path.txt  
- Check folder permissions  
- Verify PNG file creation  

---

## Summary

Start with:

=Ping()

If you see:

OK | R version ...

you are ready to go.

From there, build up:

- simple expressions  
- data transfer  
- wrapper functions  
- numerical methods  
- plotting  

This progression ensures everything is working step by step.
