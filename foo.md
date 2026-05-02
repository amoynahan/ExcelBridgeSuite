# RExcelBridge Usage Guide

This guide explains how to use RExcelBridge, the R-based add-in in ExcelBridgeSuite.

RExcelBridge is documented first because it is the primary working example in the suite. It provides the clearest reference implementation for how the bridges are intended to work: how Excel calls into a language runtime, how values are passed back and forth, how user-defined wrapper functions are organized, and how plots are generated.

JuliaExcelBridge and PythonExcelBridge follow the same general design, but they are documented separately in their own directories.

If you understand the RExcelBridge workflow first, the Julia and Python bridges will be much easier to use.

---

## Quick Start

### Step 1 — Check connectivity

=RPing()

Expected:
OK | R version ...

### Step 2 — Evaluate a simple expression

=REval("1+1")

Expected:
2

### Step 3 — Return multiple values

=REval("c(10,20,30)")

---

## Before You Start

- Add-in is attached and checked  
- rscript-path.txt points to Rscript.exe  
- Files remain in publish folder  

---

## Core Workflow

### Evaluate R code
=REval("sqrt(16)")

### Call an R function
=RCall("sum",1,2,3)

### Return structured data
=REval("matrix(c(1,2,3,4), nrow=2)")

### Pass data
=RSet("x", A1:B3)
=RGet("x")

---

## Custom Functions

Place in RFunctions.R

Example:

add_ten <- function(x) {
  x + 10
}

Excel:
=REval("add_ten(5)")

---

## Plotting

Simple:
=RPlot("plot(1:5)", "BasicPlot", 800, 600)

Dynamic:
Use RPlotDataNamed + PlotLink (two-cell pattern)

---

## Performance

See PERFORMANCE.md

---

## Troubleshooting

Ping fails → reload add-in  
R fails → check rscript-path.txt  
Plots fail → check plot-path.txt  

---

## Function Reference

RPing()
REval(code)
RCall(fun,...)
RSet(name,value)
RGet(name)
RPlot(...)
RSource(file)
