# RExcelBridge regression tests

```excel
=RPing()
```

```excel
=REval("x <- matrix(rnorm(10000*20), nrow=10000, ncol=20)")
=RGetNumeric("x")
=RLastTransfer()
```

```excel
=REval("make_matrix <- function(n, p) matrix(rnorm(n*p), nrow=n, ncol=p)")
=RCall("make_matrix", 10000, 20)
=RLastTransfer()
```

Create a spilled Excel range, then send a range that does not start in row 1:

```excel
=RANDARRAY(10000,20)
=RSet("x_from_excel", D10:W10009)
=RCall("dim", RObj("x_from_excel"))
=RLastTransfer()
```

Expected dimension: `10000, 20`.

```excel
=RSetTable("df", D10:F13, TRUE)
=RCall("names", RObj("df"))
=RGetTable("df")
=RLastTransfer()
```

```excel
=RSetTable("df_no_headers", D10:F13, FALSE)
=RCall("names", RObj("df_no_headers"))
=RGetTable("df_no_headers")
```
