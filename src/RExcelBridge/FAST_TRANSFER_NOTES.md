# RExcelBridge Fast Transfer Notes

This build keeps the existing public functions and adds fast return paths for larger objects.

## Main ideas

- `REval()` remains the general execute-code path.
- `RGet()` remains the legacy/general object getter.
- `RGetNumeric()` returns numeric vectors and matrices through a binary double buffer.
- `RGetTable()` returns data.frames/tibbles through typed column metadata.
- `RCall()` now auto-dispatches returned numeric vectors/matrices and data.frames/tibbles through the same fast paths.

## Numeric path

R writes a little-endian binary `double` buffer plus row/column metadata. C# reads the buffer into `double[]`, reshapes from R column-major order into Excel row/column cells, and returns one `object[,]` spill result.

## Table path

R exports data.frames/tibbles by column. Numeric columns use binary double buffers. Character/logical columns use JSON metadata. C# reconstructs one `object[,]` with headers.

## Compatibility

The older JSON/general path is still present for small objects, strings, lists, summaries, and other objects. Superseded code was not removed aggressively so existing functions are less likely to break during testing.

## Suggested Excel tests

```excel
=RPing()
=REval("x <- matrix(1:12, nrow=3, ncol=4)")
=RGetNumeric("x")
=RCall("MakeNumericMatrix", 10000, 20)
=RCall("MakeMixedTable", 10)
=RGetTable("some_data_frame")
```
