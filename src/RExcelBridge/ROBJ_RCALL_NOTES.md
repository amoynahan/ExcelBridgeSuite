# RObj / RCall object reference support

This build adds `RObj(name)` so `RCall()` can pass an existing R object by reference instead of passing a literal string.

## Why this is needed

Without `RObj()`, this formula:

```excel
=RCall("names", "df")
```

passes the literal string `"df"` to R. In R that is equivalent to:

```r
names("df")
```

The correct object-reference call is now:

```excel
=RCall("names", RObj("df"))
```

which is equivalent to:

```r
names(df)
```

## Tests

```excel
=RSetTable("df", D10:F13, TRUE)
=RCall("names", RObj("df"))
```

Expected output:

```text
A
B
C
```

Other useful tests:

```excel
=RCall("nrow", RObj("df"))
=RCall("ncol", RObj("df"))
=RCall("summary", RObj("df"))
```

Plain strings still work as strings:

```excel
=RCall("toupper", "abc")
```

Expected output:

```text
ABC
```
