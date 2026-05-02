# RExcelBridge Performance Guide

This document describes performance considerations and optimized workflows for RExcelBridge.

RExcelBridge is designed around a persistent R worker and JSON-based data exchange. Understanding how data moves between Excel and R is key to achieving good performance.

---

## Architecture Overview

RExcelBridge uses:

- A persistent R session (worker.R)
- JSON serialization for data transfer
- An internal object store in R for reused objects

Key implications:

- Data transfer is the main cost
- Computation inside R is fast once data is present
- Objects can persist across calls

---

## Key Performance Principles

### 1. Minimize Data Movement

The slowest step is moving data between Excel and R.

Avoid:
- Repeated large range transfers
- Recomputing the same dataset

Prefer:
- Transfer once
- Reuse in R

---

### 2. Use Persistent Objects

RExcelBridge stores objects in a persistent session.

Example:

=RSet("x", A1:A100000)
=REval("mean(x)")

Avoid:

=RCall("mean", A1:A100000)

---

### 3. Separate Transfer and Compute

Recommended pattern:

=RSet("x", A1:D100000)
=REval("result <- some_function(x)")
=RGet("result")

---

## Data Transfer Behavior

RExcelBridge converts data using JSON.

Important characteristics:

- Data frames are converted to row-wise lists
- Matrices are transferred as row arrays
- Scalars are returned directly
- Large objects increase serialization cost

---

## Internal Object Store

The worker maintains an internal object store.

Benefits:

- Objects persist across calls
- Reuse avoids repeated transfers
- Enables multi-step workflows

Use:

=RRemove("x")

to clean up memory.

---

## Transfer Diagnostics

RExcelBridge tracks the last transfer using:

.last_transfer_info

This includes:

- Method (RSet, RGet, etc.)
- Object name
- Type and class
- Dimensions
- Rows and columns
- Elapsed time

This can be used for performance debugging.

---

## Large Data Recommendations

Use optimized patterns when:

- Data > 10,000 rows
- Repeated calculations
- Multi-step workflows

Best practice:

1. Load once with RSet
2. Transform inside R
3. Return only final results

---

## Recalculation Strategy

Excel recalculation can trigger expensive operations.

Recommendations:

- Avoid volatile dependencies
- Use helper cells
- Trigger recalculation manually (F9) for large jobs

---

## Plot Performance

### Simple plotting

Use RPlot for one-time rendering.

### Dynamic plotting

Use two-step approach:

- RPlotDataNamed → generate plot
- PlotLink → display image

This prevents unnecessary recomputation.

---

## Memory Management

Objects persist in R until removed.

Recommendations:

- Remove unused objects
- Avoid duplicating large objects
- Reuse existing data where possible

---

## When Performance Matters

Use these strategies when:

- Working with large datasets
- Repeatedly recalculating
- Experiencing slow response times

---

## Summary

For best performance:

- Minimize data transfers
- Use persistent R objects
- Separate data transfer from computation
- Avoid repeated large RCall operations
- Use dynamic plotting patterns carefully

---

## Next Steps

You can extend this document with:

- Benchmarks (small vs large data)
- Timing comparisons
- Screenshots of workflows

