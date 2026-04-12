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
