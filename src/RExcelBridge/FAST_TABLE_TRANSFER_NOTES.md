# Fast data.frame / tibble transfer tests

This build adds `RSetTable()` for Excel -> R data.frame transfer and keeps `RGetTable()` / `RCall()` typed table returns.

## R -> Excel data.frame / tibble

```excel
=REval("df <- data.frame(id=c('A','B','C'), x=c(1.1,2.2,3.3), group=c('low','mid','high'), stringsAsFactors=FALSE)")
=RGetTable("df")
=RLastTransfer()
```

Expected: headers plus 3 rows. `x` returns through the numeric column path.

## RCall auto-dispatch table return

```excel
=REval("make_df <- function(n) data.frame(id=paste0('ID',1:n), x=rnorm(n), y=runif(n), group=sample(c('A','B'), n, TRUE), stringsAsFactors=FALSE)")
=RCall("make_df", 1000)
=RLastTransfer()
```

Expected: method = typed table.

## Excel -> R data.frame

Create a table that starts away from row 1, for example:

```excel
D10: id      E10: x          F10: group
D11: A       E11: 1.1        F11: low
D12: B       E12: 2.2        F12: mid
D13: C       E13: 3.3        F13: high
```

Then run:

```excel
=RSetTable("df_from_excel", D10:F13, TRUE)
=RCall("str", "df_from_excel")
=RGetTable("df_from_excel")
=RLastTransfer()
```

Expected: data.frame with character `id`, numeric `x`, and character `group`.

## Large mixed table stress test

Use Excel formulas to create columns such as ID, numeric values, and category, then:

```excel
=RSetTable("bigdf", D10:G10009, TRUE)
=RCall("dim", "bigdf")
=RGetTable("bigdf")
```
