# JuliaExcelBridge fast-transfer test notes

Quick build/runtime smoke tests:

```excel
=JPing()
=JEval("x = randn(10000,20)")
=JGetNumeric("x")
=JLastTransfer()
```

Object-reference test:

```excel
=JCall("size", JObj("x"))
```

Excel to Julia numeric transfer:

```excel
=RANDARRAY(10000,20)
=JSet("y", D10:W10009)
=JCall("size", JObj("y"))
=JLastTransfer()
```

DataFrame path requires DataFrames.jl in Julia:

```julia
using Pkg
Pkg.add("DataFrames")
```

Then test:

```excel
=JEval("using DataFrames; df = DataFrame(id=["A","B","C"], x=[1.1,2.2,3.3], group=["low","mid","high"])")
=JGetTable("df")
=JLastTransfer()
```

Excel to Julia DataFrame:

```excel
=JSetTable("df2", D10:F13, TRUE)
=JGetTable("df2")
=JCall("names", JObj("df2"))
```
