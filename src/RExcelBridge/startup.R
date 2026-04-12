suppressPackageStartupMessages({
  library(jsonlite)
  library(flexsurv)
  library(survival)
})

hello_world <- function() {
  "Hello from R"
}

add_numbers <- function(x, y) {
  x + y
}

make_seq_table <- function(from, to, by = 1) {
  data.frame(
    x = seq(from, to, by = by),
    y = seq(from, to, by = by)^2
  )
}

fit_weibull_simple <- function(time, event) {
  dat <- data.frame(
    time = as.numeric(unlist(time)),
    event = as.numeric(unlist(event))
  )

  flexsurvreg(
    survival::Surv(time, event) ~ 1,
    data = dat,
    dist = "weibull"
  )
}

model_summary_table <- function(handle) {
  obj <- get(handle, envir = .object_store)

  if (inherits(obj, "flexsurvreg")) {
    est <- as.data.frame(obj$res[, c("est", "L95%", "U95%"), drop = FALSE])
    est$parameter <- rownames(est)
    est <- est[, c("parameter", "est", "L95%", "U95%")]
    rownames(est) <- NULL
    return(est)
  }

  stop("Unsupported model class for summary extraction.")
}

surv_table <- function(handle, times) {
  obj <- get(handle, envir = .object_store)
  times <- as.numeric(unlist(times))

  s <- summary(obj, t = times, type = "survival")[[1]]

  data.frame(
    time = s$time,
    surv = s$est
  )
}

drop_object <- function(handle) {
  if (exists(handle, envir = .object_store, inherits = FALSE)) {
    rm(list = handle, envir = .object_store)
    return(TRUE)
  }
  FALSE
}

source("RFunctions.R")
