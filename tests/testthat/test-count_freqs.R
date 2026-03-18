library(testthat)

# Setup: penguins from datasets (available since R 4.1), with simulated weights
# Weights are normalized so mean = 1 and sum = nrow(penguins)
penguins_data <- datasets::penguins
n_rows <- nrow(penguins_data)

set.seed(42)
weights_raw <- runif(n_rows, 0.5, 1.5)
penguins_data$w <- weights_raw / mean(weights_raw)

# 1. When weight = NULL, sum of n equals nrow --------------------------------

test_that("sum of n equals nrow when weight = NULL", {
  result <- suppressWarnings(
    count_freqs(penguins_data, var = "species", weight = NULL, prop = "none")
  )
  expect_equal(sum(result$n), n_rows)
})

# 2. When weights specified, sum of n equals sum of weights ------------------

test_that("sum of n equals sum of weights when weights are specified", {
  result <- count_freqs(
    penguins_data, var = "species", weight = "w", prop = "none"
  )
  expect_equal(sum(result$n), sum(penguins_data$w))
})

# 3. prop = "total", group = NULL: total column sums to 1 --------------------

test_that("total column sums to 1 when prop = 'total' and group = NULL", {
  result <- suppressWarnings(
    count_freqs(penguins_data, var = "species", weight = NULL, prop = "total")
  )
  expect_equal(sum(result$total), 1)
})

# 4. prop = "total", group not NULL: all numeric columns combined sum to 1 ---

test_that("all numeric columns combined sum to 1 when prop = 'total' and group is not NULL", {
  result <- count_freqs(
    penguins_data,
    var    = "species",
    group  = "island",
    weight = "w",
    prop   = "total"
  )
  numeric_cols <- names(result)[sapply(result, is.numeric)]
  num_df <- as.data.frame(result)[, numeric_cols, drop = FALSE]
  expect_equal(sum(colSums(num_df)), 1, tolerance = 1e-10)
})

# 5. prop = "row", group not NULL: each row of numeric columns sums to 1 -----

test_that("row sums of numeric columns equal 1 when prop = 'row' and group is not NULL", {
  result <- count_freqs(
    penguins_data,
    var    = "species",
    group  = "island",
    weight = "w",
    prop   = "row"
  )
  numeric_cols <- names(result)[sapply(result, is.numeric)]
  num_df <- as.data.frame(result)[, numeric_cols, drop = FALSE]
  expect_equal(rowSums(num_df), rep(1, nrow(result)), tolerance = 1e-10)
})

# 6. prop = "col", group not NULL: each numeric column sums to 1 -------------

test_that("each numeric column sums to 1 when prop = 'col' and group is not NULL", {
  result <- count_freqs(
    penguins_data,
    var    = "species",
    group  = "island",
    weight = "w",
    prop   = "col"
  )
  numeric_cols <- names(result)[sapply(result, is.numeric)]
  num_df <- as.data.frame(result)[, numeric_cols, drop = FALSE]
  expected <- rep(1, length(numeric_cols))
  names(expected) <- numeric_cols
  expect_equal(colSums(num_df), expected, tolerance = 1e-10)
})

# 7. include_empty = TRUE: empty factor levels appear with n = 0 --------------

test_that("include_empty = TRUE includes empty factor levels with n = 0", {
  d <- data.frame(
    x = factor(c("a", "a", "b"), levels = c("a", "b", "c")),
    w = c(1, 1, 1)
  )
  result <- count_freqs(d, var = "x", weight = "w", prop = "none", include_empty = TRUE)
  expect_true("c" %in% result$x)
  expect_equal(result[result$x == "c", n], 0)
})

test_that("include_empty = FALSE excludes empty factor levels", {
  d <- data.frame(
    x = factor(c("a", "a", "b"), levels = c("a", "b", "c")),
    w = c(1, 1, 1)
  )
  result <- count_freqs(d, var = "x", weight = "w", prop = "none", include_empty = FALSE)
  expect_false("c" %in% result$x)
})

test_that("include_empty = TRUE works with grouping var", {
  d <- data.frame(
    species = factor(c("a", "a", "b"), levels = c("a", "b", "c")),
    island  = factor(c("p", "q", "p"), levels = c("p", "q")),
    w       = c(1, 1, 1)
  )
  result <- count_freqs(d, var = "species", group = "island", weight = "w", prop = "none", include_empty = TRUE)
  # "c" should appear as a row with 0 counts in both group columns
  expect_true("c" %in% result$species)
  row_c <- as.data.frame(result)[result$species == "c", c("p", "q")]
  expect_equal(unlist(row_c), c(p = 0, q = 0))
})
