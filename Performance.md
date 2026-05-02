# RExcelBridge Performance Guide

This document explains how to use RExcelBridge efficiently when moving data between Excel and R.

RExcelBridge supports simple general-purpose transfer functions, but it also includes more targeted transfer functions for tables and numeric data. For larger datasets, these targeted functions should usually be preferred.

---

## Overview

The main performance cost in RExcelBridge is moving data between Excel and R.

Once data is inside the persistent R session, R computations are usually fast. The best performance strategy is therefore:

1. Move data into R efficiently
2. Keep it in the persistent R session
3. Run computations inside R
4. Return only the results needed in Excel

---

## Recommended Transfer Functions

For performance-sensitive work, focus on these functions:

| Function | Direction | Best Use |
|---|---|---|
| `RSetTable(name, range, hasHeaders)` | Excel to R | Send rectangular Excel data into R as a table/data frame |
| `RGetTable(name)` | R to Excel | Return an R data frame/table to Excel |
| `RGetNumeric(name)` | R to Excel | Return a numeric vector or matrix using the numeric fast path |

These functions make the data-transfer intent explicit and help avoid unnecessary overhead.

---

## General-Purpose vs Performance-Oriented Transfer

RExcelBridge also provides general-purpose functions:

| Function | Direction | Use |
|---|---|---|
| `RSet(name, value)` | Excel to R | General assignment of Excel values/ranges to R |
| `RGet(name)` | R to Excel | General retrieval of R objects |
| `RCall(fun, ...)` | Excel to R and back | Call an R function with Excel arguments |
| `REval(code)` | R execution | Execute R code and optionally return a result |

These are convenient and useful for small examples.

For larger data, prefer the more explicit table and numeric functions where appropriate.

---

## RSetTable: Send Excel Tables to R

Use `RSetTable` when sending rectangular Excel data into R.

This is the preferred approach for worksheet data that should become an R data frame.

Example:

```excel
=RSetTable("df", A1:D100000, TRUE)
```

The third argument indicates whether the first row contains headers.

Use `TRUE` when the first row contains column names:

```excel
=RSetTable("df", A1:D100000, TRUE)
```

Use `FALSE` when the range contains data only:

```excel
=RSetTable("df", A1:D99999, FALSE)
```

After the table is loaded, work with it inside R:

```excel
=REval("summary(df)")
```

or:

```excel
=REval("result <- aggregate(df$VALUE, list(df$GROUP), mean)")
```

---

## RGetTable: Return Data Frames to Excel

Use `RGetTable` when returning a data frame or rectangular table from R to Excel.

Example:

```excel
=RGetTable("result")
```

This is preferred when the R object is a data frame or table-like object.

Recommended pattern:

```excel
=REval("result <- head(df, 20)")
=RGetTable("result")
```

This keeps the larger object in R and only returns the table you need.

---

## RGetNumeric: Fast Numeric Return

Use `RGetNumeric` when returning numeric vectors or matrices from R to Excel.

Example:

```excel
=REval("m <- matrix(rnorm(100000), ncol=10)")
=RGetNumeric("m")
```

This is the preferred return path for large numeric matrices.

Use it for:

- Numeric matrices
- Numeric vectors
- Simulation output
- Model matrices
- Numeric result arrays

Avoid using it for:

- Mixed-type data frames
- Character columns
- Tables with headers
- Complex nested objects

For mixed or table-like data, use `RGetTable`.

---

## Recommended Large-Data Workflow

### Step 1 — Load data once

```excel
=RSetTable("df", A1:D100000, TRUE)
```

### Step 2 — Compute inside R

```excel
=REval("result <- subset(df, GROUP == 'A')")
```

### Step 3 — Return only the needed result

```excel
=RGetTable("result")
```

For numeric outputs:

```excel
=REval("m <- as.matrix(df[, c('X1','X2','X3')])")
=RGetNumeric("m")
```

---

## Avoid Repeated Large RCall Transfers

This pattern is convenient but can be inefficient for large ranges:

```excel
=RCall("some_function", A1:D100000)
```

The range may be transferred again each time Excel recalculates.

Prefer:

```excel
=RSetTable("df", A1:D100000, TRUE)
=REval("result <- some_function(df)")
=RGetTable("result")
```

This separates data transfer from computation.

---

## Choosing the Right Function

### Use RSetTable when:

- The input is a rectangular Excel range
- The data should become an R data frame
- The first row may contain column names
- The dataset has mixed column types

### Use RGetTable when:

- The R object is a data frame
- You want headers returned to Excel
- The output includes mixed types
- The result is table-like

### Use RGetNumeric when:

- The R object is numeric
- The output is a vector or matrix
- You want the fastest numeric return path
- You do not need column headers

### Use RSet / RGet when:

- The object is small
- You are testing
- You want the simplest general-purpose behavior

---

## Persistent R Session

RExcelBridge uses a persistent R worker.

This means objects remain available after they are created.

Example:

```excel
=RSetTable("df", A1:D100000, TRUE)
=REval("nrow(df)")
=REval("names(df)")
=REval("summary(df)")
```

The data does not need to be resent for each command.

---

## Memory Management

Because objects persist in R, large objects remain in memory until removed.

Use:

```excel
=RRemove("df")
```

to remove objects that are no longer needed.

You can inspect the session with:

```excel
=RObjects()
```

and describe an object with:

```excel
=RDescribe("df")
```

---

## Data Frame Support

RExcelBridge supports standard rectangular data frames.

Supported:

- Numeric columns
- Character columns
- Logical columns
- Standard rectangular tables

Avoid:

- List columns
- Nested data frames
- Highly complex R objects

For best results, keep data frames rectangular and Excel-like.

---

## Numeric Fast Path

`RGetNumeric` is intended for numeric-only output.

This is useful because numeric matrices can be returned more efficiently than mixed-type tables.

Good candidates:

```r
matrix(rnorm(100000), ncol = 10)
as.matrix(df[, numeric_columns])
predict(model, newdata)
```

If the output is mixed type or requires headers, use `RGetTable`.

---

## Plot Performance

Plotting should follow the same principle: avoid unnecessary transfers.

### Simple plots

Use `RPlot` when creating a one-time plot.

### Dynamic plots

Use the two-cell pattern:

1. `RPlotDataNamed` creates the plot and returns the PNG path
2. `PlotLink` displays the image in Excel

This separates plot generation from image display and improves reliability during recalculation.

---

## Recalculation Strategy

Excel may recalculate formulas more often than expected.

For large workflows:

- Load large data once
- Use `REval` to compute inside R
- Return only small or final results
- Consider manual recalculation with `F9`
- Avoid volatile formulas connected to large transfers

---

## Suggested Performance Pattern

For table workflows:

```excel
=RSetTable("df", A1:D100000, TRUE)
=REval("result <- transform(df, NEW_VALUE = VALUE * 2)")
=RGetTable("result")
```

For numeric workflows:

```excel
=REval("m <- matrix(rnorm(100000), ncol=10)")
=RGetNumeric("m")
```

For cleanup:

```excel
=RRemove("df")
=RRemove("result")
=RRemove("m")
```

---

## Summary

For best performance:

- Use `RSetTable` to send Excel tables into R
- Use `RGetTable` to return data frames and mixed-type tables
- Use `RGetNumeric` to return numeric vectors and matrices
- Avoid repeated large `RCall` transfers
- Keep large objects in the persistent R session
- Return only the results needed in Excel
- Remove large objects when finished

---

## Suggested Benchmarks to Add Later

You may want to add benchmark examples such as:

- `RGet` vs `RGetNumeric` for numeric matrices
- `RSet` vs `RSetTable` for rectangular data
- Repeated `RCall` vs `RSetTable` + `REval`
- 10,000 rows vs 100,000 rows vs 200,000 rows

These benchmarks would make the performance benefits more concrete.
