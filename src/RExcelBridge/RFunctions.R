CholDecomp <- function(mat) {
  m <- as.matrix(mat)
  storage.mode(m) <- "double"
  chol(m)
}

CholDecompLower <- function(mat) {
  m <- as.matrix(mat)
  storage.mode(m) <- "double"
  t(chol(m))
}

ShowObjectInfo <- function(x) {
  cls <- paste(class(x), collapse = ",")
  dm <- dim(x)
  dim_text <- if (is.null(dm)) "NULL" else paste(dm, collapse = "x")
  paste("class =", cls, "; dim =", dim_text, "; length =", length(x))
}

# Fast-path test helpers ----------------------------------------------------
MakeNumericMatrix <- function(rows = 1000, cols = 20) {
  matrix(stats::rnorm(as.integer(rows) * as.integer(cols)), nrow = as.integer(rows), ncol = as.integer(cols))
}

MakeMixedTable <- function(rows = 10) {
  rows <- as.integer(rows)
  data.frame(
    id = sprintf("ID%05d", seq_len(rows)),
    value = seq_len(rows) / 10,
    flag = seq_len(rows) %% 2 == 0,
    stringsAsFactors = FALSE
  )
}
