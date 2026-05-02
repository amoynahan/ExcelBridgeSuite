# RExcelBridge Usage Guide

This guide explains how to use RExcelBridge, the R-based add-in in ExcelBridgeSuite.

RExcelBridge is documented first because it is the primary working example in the suite. It provides the clearest reference implementation for how the bridges are intended to work: how Excel calls into a language runtime, how values are passed back and forth, how user-defined wrapper functions are organized, and how plots are generated.

JuliaExcelBridge and PythonExcelBridge follow the same general design, but they are documented separately in their own directories.

If you understand the RExcelBridge workflow first, the Julia and Python bridges will be much easier to use.

---

## Quick Start

Start here. This verifies everything is working.

### Step 1 — Check connectivity

```excel
=RPing()
