# RExcelBridge fast transfer next step

This version keeps the working fast numeric transfer path and adds the next step from the design discussion.

## What changed

1. `RCall()` now uses return-type dispatch automatically.
   - Numeric scalar/vector/matrix -> fast binary numeric path.
   - `data.frame` / tibble -> typed table path.
   - Other small/general objects -> existing JSON/general path.

2. `RGetNumeric()` is retained as the explicit/debug fast numeric getter.

3. `RGetTable()` is retained as the explicit/debug data-frame getter.

4. Added `RLastTransfer()`.
   - Shows the last transfer method used.
   - Shows source, name, class, type, dimensions, rows, columns, and R-side export time.

5. Improved C# binary double reading.
   - Replaced per-double `BinaryReader.ReadDouble()` loop with a block byte read plus `Buffer.BlockCopy()`.
   - This should reduce overhead for large numeric transfers.

## Suggested tests

```excel
=RPing()
```

```excel
=REval("make_matrix <- function(n, p) matrix(rnorm(n*p), nrow=n, ncol=p)")
=RCall("make_matrix", 10000, 20)
=RLastTransfer()
```

Expected transfer method:

```text
fast numeric binary
```

Data frame test:

```excel
=REval("make_df <- function(n) data.frame(id=paste0('id',1:n), x=rnorm(n), y=runif(n))")
=RCall("make_df", 10)
=RLastTransfer()
```

Expected transfer method:

```text
typed table
```

General object test:

```excel
=RCall("paste", "hello", "world")
=RLastTransfer()
```

Expected transfer method:

```text
json/general
```

## Notes

The old general path is still present. The new behavior is designed to be additive and backward-compatible.
