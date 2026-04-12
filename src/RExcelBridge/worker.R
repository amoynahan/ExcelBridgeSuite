args <- commandArgs(trailingOnly = TRUE)
startup_file <- if (length(args) >= 1) args[[1]] else NULL

options(warn = 1)

suppressPackageStartupMessages(library(jsonlite))

`%||%` <- function(x, y) if (is.null(x)) y else x

.object_store <- new.env(parent = emptyenv())

if (!is.null(startup_file) && file.exists(startup_file)) {
  source(startup_file, local = .GlobalEnv)
}

resolve_function <- function(fun_name) {
  if (grepl(":::", fun_name, fixed = TRUE)) {
    parts <- strsplit(fun_name, ":::", fixed = TRUE)[[1]]
    pkg <- parts[1]
    fn  <- parts[2]
    return(get(fn, envir = asNamespace(pkg), inherits = FALSE))
  }

  if (grepl("::", fun_name, fixed = TRUE)) {
    parts <- strsplit(fun_name, "::", fixed = TRUE)[[1]]
    pkg <- parts[1]
    fn  <- parts[2]
    return(getExportedValue(pkg, fn))
  }

  get(fun_name, envir = .GlobalEnv, inherits = TRUE)
}

coerce_for_json <- function(x, include_headers = TRUE) {
  if (is.null(x)) return(NULL)

  if (is.data.frame(x)) {
    bad_cols <- vapply(x, function(col) is.list(col) && !is.atomic(col), logical(1))
    if (any(bad_cols)) {
      stop("Data frames with list-columns are not supported.")
    }

    x[] <- lapply(x, function(col) {
      if (is.factor(col)) as.character(col) else col
    })

    header_row <- as.list(unname(names(x)))
    body_rows <- lapply(seq_len(nrow(x)), function(i) {
      as.list(unname(as.matrix(x[i, , drop = FALSE])[1, ]))
    })

    if (isTRUE(include_headers)) {
      return(unname(c(list(header_row), body_rows)))
    }

    return(unname(body_rows))
  }

  if (is.matrix(x)) {
    return(unname(lapply(seq_len(nrow(x)), function(i) as.list(unname(x[i, ])))))
  }

  if (is.atomic(x) && is.null(dim(x))) {
    if (length(x) == 1) {
      return(x[[1]])
    } else {
      return(as.list(unname(x)))
    }
  }

  if (is.list(x)) {
    return(unname(x))
  }

  as.character(x)
}

make_handle <- function(prefix = "obj") {
  paste0(prefix, "_", format(Sys.time(), "%Y%m%d_%H%M%S"), "_", sample.int(1e6, 1))
}

safe_eval <- function(expr_fun) {
  tryCatch(
    list(ok = TRUE, result = expr_fun()),
    error = function(e) list(ok = FALSE, error = conditionMessage(e))
  )
}

is_scalar_json_value <- function(x) {
  is.null(x) || is.atomic(x)
}

is_json_matrix_like <- function(x) {
  is.list(x) &&
    length(x) > 0 &&
    all(vapply(x, is.list, logical(1))) &&
    all(vapply(x, function(row) length(row) > 0, logical(1))) &&
    length(unique(vapply(x, length, integer(1)))) == 1 &&
    all(vapply(x, function(row) all(vapply(row, is_scalar_json_value, logical(1))), logical(1)))
}

common_atomic_mode <- function(x) {
  if (all(vapply(x, function(v) is.null(v) || is.logical(v), logical(1)))) return("logical")
  if (all(vapply(x, function(v) is.null(v) || is.numeric(v), logical(1)))) return("numeric")
  if (all(vapply(x, function(v) is.null(v) || is.character(v), logical(1)))) return("character")
  "character"
}

coerce_atomic_vector <- function(x) {
  mode <- common_atomic_mode(x)
  if (identical(mode, "logical")) return(unlist(lapply(x, as.logical), use.names = FALSE))
  if (identical(mode, "numeric")) return(unlist(lapply(x, as.numeric), use.names = FALSE))
  unlist(lapply(x, function(v) if (is.null(v)) NA_character_ else as.character(v)), use.names = FALSE)
}

json_matrix_to_r <- function(x) {
  nr <- length(x)
  nc <- length(x[[1]])
  flat <- unlist(x, recursive = TRUE, use.names = FALSE)
  vec <- coerce_atomic_vector(as.list(flat))
  matrix(vec, nrow = nr, ncol = nc, byrow = TRUE)
}

convert_json_value <- function(x) {
  if (is.null(x)) return(NULL)
  if (is.atomic(x)) return(x)
  if (!is.list(x)) return(x)
  if (length(x) == 0) return(list())

  if (is_json_matrix_like(x)) {
    return(json_matrix_to_r(x))
  }

  converted <- lapply(x, convert_json_value)

  if (all(vapply(converted, is_scalar_json_value, logical(1)))) {
    return(coerce_atomic_vector(converted))
  }

  converted
}

convert_call_arg <- function(x) {
  convert_json_value(x)
}

convert_assignment_value <- function(x) {
  convert_json_value(x)
}

normalize_call_args <- function(args) {
  if (is.null(args)) {
    return(list())
  }

  if (!is.list(args)) {
    return(list(convert_call_arg(args)))
  }

  lapply(args, convert_call_arg)
}

format_dim <- function(x) {
  d <- dim(x)
  if (is.null(d)) {
    if (length(x) == 1) return("1 x 1")
    return(paste0(length(x), " x 1"))
  }
  paste(d, collapse = " x ")
}

object_summary_row <- function(name, x) {
  data.frame(
    Name = name,
    Class = paste(class(x), collapse = ", "),
    Type = typeof(x),
    Length = as.character(length(x)),
    Dimensions = format_dim(x),
    stringsAsFactors = FALSE
  )
}

object_describe_table <- function(name, x) {
  data.frame(
    Field = c("Name", "Class", "Type", "Length", "Dimensions"),
    Value = c(
      name,
      paste(class(x), collapse = ", "),
      typeof(x),
      as.character(length(x)),
      format_dim(x)
    ),
    stringsAsFactors = FALSE
  )
}

is_assignment_call <- function(expr) {
  is.call(expr) && identical(as.character(expr[[1]]), "<-")
}

assignment_target_name <- function(expr) {
  if (!is_assignment_call(expr)) return(NULL)
  if (length(expr) < 2) return(NULL)

  lhs <- expr[[2]]
  if (is.symbol(lhs)) {
    return(as.character(lhs))
  }

  NULL
}



render_plot_to_file <- function(code, file, width = 800, height = 600, res = 96) {
  if (is.null(file) || !nzchar(file)) {
    stop("Plot file path is blank.")
  }

  dir.create(dirname(file), recursive = TRUE, showWarnings = FALSE)

  grDevices::png(filename = file, width = width, height = height, res = res)
  on.exit(grDevices::dev.off(), add = TRUE)

  result <- eval(parse(text = code), envir = .GlobalEnv)
  if (inherits(result, "ggplot")) {
    print(result)
  }

  normalizePath(file, winslash = "/", mustWork = FALSE)
}


eval_code_for_excel <- function(code) {
  exprs <- parse(text = code)
  if (length(exprs) == 0) return(NULL)

  last_value <- NULL
  assigned_name <- NULL

  for (expr in exprs) {
    target_name <- assignment_target_name(expr)
    last_value <- eval(expr, envir = .GlobalEnv)
    if (!is.null(target_name)) {
      assigned_name <- target_name
    }
  }

  if (!is.null(assigned_name)) {
    return(sprintf("Assigned: %s", assigned_name))
  }

  coerce_for_json(last_value)
}

handle_request <- function(req) {
  cmd <- req$cmd %||% ""

  if (identical(cmd, "ping")) {
    return(list(id = req$id, ok = TRUE, result = paste("OK |", R.version.string)))
  }

  if (identical(cmd, "source")) {
    out <- safe_eval(function() {
      source(req$file, local = .GlobalEnv)
      TRUE
    })
    return(c(list(id = req$id), out))
  }

  if (identical(cmd, "eval")) {
    out <- safe_eval(function() {
      eval_code_for_excel(req$code)
    })
    return(c(list(id = req$id), out))
  }

  if (identical(cmd, "plot")) {
    out <- safe_eval(function() {
      render_plot_to_file(
        code = req$code,
        file = req$file,
        width = as.integer(req$width %||% 800),
        height = as.integer(req$height %||% 600),
        res = as.integer(req$res %||% 96)
      )
    })
    return(c(list(id = req$id), out))
  }

  if (identical(cmd, "call")) {
    out <- safe_eval(function() {
      f <- resolve_function(req$fun)
      call_args <- normalize_call_args(req$args %||% list())
      value <- do.call(f, call_args)

      if (inherits(value, c("flexsurvreg", "coxph", "survfit", "lm", "glm"))) {
        handle <- make_handle("fit")
        assign(handle, value, envir = .object_store)
        return(list(handle = handle, class = class(value)[1]))
      }

      coerce_for_json(value)
    })
    return(c(list(id = req$id), out))
  }

  if (identical(cmd, "set")) {
    out <- safe_eval(function() {
      value <- convert_assignment_value(req$value)
      assign(req$name, value, envir = .GlobalEnv)
      TRUE
    })
    return(c(list(id = req$id), out))
  }

  if (identical(cmd, "get")) {
    out <- safe_eval(function() {
      if (!exists(req$name, envir = .GlobalEnv, inherits = FALSE)) {
        stop(sprintf("Object '%s' was not found.", req$name))
      }
      value <- get(req$name, envir = .GlobalEnv, inherits = FALSE)
      coerce_for_json(value)
    })
    return(c(list(id = req$id), out))
  }

  if (identical(cmd, "exists")) {
    out <- safe_eval(function() {
      exists(req$name, envir = .GlobalEnv, inherits = FALSE)
    })
    return(c(list(id = req$id), out))
  }

  if (identical(cmd, "remove")) {
    out <- safe_eval(function() {
      existed <- exists(req$name, envir = .GlobalEnv, inherits = FALSE)
      if (existed) {
        rm(list = req$name, envir = .GlobalEnv)
      }
      existed
    })
    return(c(list(id = req$id), out))
  }

  if (identical(cmd, "objects")) {
    out <- safe_eval(function() {
      nms <- ls(envir = .GlobalEnv, all.names = FALSE)
      if (length(nms) == 0) {
        return(data.frame(
          Name = character(),
          Class = character(),
          Type = character(),
          Length = character(),
          Dimensions = character(),
          stringsAsFactors = FALSE
        ))
      }
      rows <- lapply(nms, function(nm) object_summary_row(nm, get(nm, envir = .GlobalEnv, inherits = FALSE)))
      coerce_for_json(do.call(rbind, rows))
    })
    return(c(list(id = req$id), out))
  }

  if (identical(cmd, "describe")) {
    out <- safe_eval(function() {
      if (!exists(req$name, envir = .GlobalEnv, inherits = FALSE)) {
        stop(sprintf("Object '%s' was not found.", req$name))
      }
      value <- get(req$name, envir = .GlobalEnv, inherits = FALSE)
      coerce_for_json(object_describe_table(req$name, value))
    })
    return(c(list(id = req$id), out))
  }

  list(id = req$id, ok = FALSE, error = paste("Unknown command:", cmd))
}

con <- file("stdin", open = "r")

repeat {
  line <- readLines(con, n = 1, warn = FALSE)
  if (length(line) == 0)
    break

  line <- trimws(line)
  if (!nzchar(line))
    next

  req <- tryCatch(
    fromJSON(line, simplifyVector = FALSE),
    error = function(e) NULL
  )

  if (is.null(req)) {
    cat(toJSON(
      list(id = NA, ok = FALSE, error = "Invalid JSON"),
      auto_unbox = TRUE,
      null = "null"
    ), "\n")
    flush.console()
    next
  }

  resp <- handle_request(req)
  cat(toJSON(resp, auto_unbox = TRUE, null = "null"), "\n")
  flush.console()
}
