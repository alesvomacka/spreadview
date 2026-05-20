library(testthat)

df <- data.frame(
  x = factor(c("a", "b", "c")),
  y = factor(c("d", "e", "f")),
  z = 1:3
)

# check_input ------------------------------------------------------------------

test_that("check_input passes for valid inputs", {
  expect_true(spreadview:::check_input(df, vars = "x", group = "y"))
})

test_that("check_input errors when data is not a data frame", {
  expect_error(spreadview:::check_input(list(x = 1), vars = "x", group = character(0)),
               "'data' must be a data frame")
})

test_that("check_input errors when vars not in data", {
  expect_error(spreadview:::check_input(df, vars = "missing", group = character(0)),
               "not found in data")
})

test_that("check_input errors when vars are not factors", {
  expect_error(spreadview:::check_input(df, vars = "z", group = character(0)),
               "non-factor")
})

test_that("check_input errors when group not in data", {
  expect_error(spreadview:::check_input(df, vars = "x", group = "missing"),
               "not found in data")
})

test_that("check_input errors when group variable is not a factor", {
  expect_error(spreadview:::check_input(df, vars = "x", group = "z"),
               "non-factor")
})

# check_weights ----------------------------------------------------------------

wdf <- data.frame(x = 1:10, w = rep(1, 10))

test_that("check_weights passes for valid weights", {
  expect_silent(spreadview:::check_weights(wdf, "w"))
})

test_that("check_weights errors when weight column missing", {
  expect_error(spreadview:::check_weights(wdf, "no_such_col"))
})

test_that("check_weights errors when weight is not numeric", {
  wdf$w <- as.character(wdf$w)
  expect_error(spreadview:::check_weights(wdf, "w"), "numeric")
})

test_that("check_weights errors when weight contains NA", {
  wdf$w[1] <- NA
  expect_error(spreadview:::check_weights(wdf, "w"), "missing")
})

test_that("check_weights errors when weight contains negative values", {
  wdf$w[1] <- -1
  expect_error(spreadview:::check_weights(wdf, "w"), "non-negative")
})
