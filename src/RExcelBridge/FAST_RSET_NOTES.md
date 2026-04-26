# Fast RSet numeric transfer

This version adds the fast numeric path for Excel-to-R assignment.

## What changed

`RSet(name, value)` now auto-detects numeric Excel inputs. If the input is a scalar, vector, or rectangular numeric range, C# writes a binary double buffer and the R worker reads it with `readBin()`.

Mixed text/numeric inputs continue to use the existing general JSON path.

## Numeric path

Excel range -> C# object[,] -> row-major double buffer -> R `readBin()` -> `matrix(..., byrow=TRUE)` -> assigned object.

## Suggested tests

```excel
=RSet("x", A1:T10000)
=RCall("dim", "x")
=RLastTransfer()
=RCall("mean", "x")
```

Small test:

```excel
=RSet("small", A1:C3)
=RGetNumeric("small")
=RLastTransfer()
```

Mixed table fallback test:

```excel
=RSet("mixed", A1:D10)
=RLastTransfer()
```

If the range contains text or Excel errors, `RSet` falls back to the general path. Blank cells inside an otherwise numeric range are transferred as `NaN`.
