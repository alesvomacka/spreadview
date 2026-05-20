library(testthat)

penguins_data <- na.omit(datasets::penguins)

# 1. No grouping: basic structure -----------------------------------------------

test_that("compose_table returns expected columns when no group", {
  result <- suppressWarnings(
    compose_table(penguins_data, var = "species", prop = "col")
  )
  expect_true(all(c("item", "label", "var", "N", "n", "total") %in% names(result)))
})

test_that("compose_table proportions sum to 1 when prop = col and no group", {
  result <- suppressWarnings(
    compose_table(penguins_data, var = "species", prop = "col")
  )
  expect_equal(sum(result$total, na.rm = TRUE), 1, tolerance = 1e-6)
})

test_that("compose_table sets item and label only on first row", {
  result <- suppressWarnings(
    compose_table(penguins_data, var = "species")
  )
  expect_equal(result$item[1], "species")
  expect_true(all(result$item[-1] == "" | is.na(result$item[-1])))
})

# 2. With grouping --------------------------------------------------------------

test_that("compose_table with group produces grouped columns", {
  result <- suppressWarnings(
    compose_table(penguins_data, var = "species", group = "island")
  )
  group_cols <- grep("^\\[island\\]", names(result), value = TRUE)
  expect_gt(length(group_cols), 0)
})

test_that("compose_table with group ends with a totals row (item and var both NA)", {
  result <- suppressWarnings(
    compose_table(penguins_data, var = "species", group = "island")
  )
  last_row <- result[nrow(result), ]
  expect_true(is.na(last_row$item))
  expect_true(is.na(last_row$var))
})

test_that("compose_table col proportions sum to 1 per group column", {
  result <- suppressWarnings(
    compose_table(penguins_data, var = "species", group = "island", prop = "col")
  )
  group_cols <- grep("^\\[island\\]", names(result), value = TRUE)
  # Exclude totals row (NA var) before summing
  data_rows <- result[!is.na(result$var), ]
  col_sums <- sapply(data_rows[, ..group_cols], sum, na.rm = TRUE)
  expect_true(all(abs(col_sums - 1) < 1e-6 | col_sums == 0))
})

# 3. Weights --------------------------------------------------------------------

test_that("compose_table weighted n sum equals sum of weights", {
  set.seed(1)
  penguins_data$w <- runif(nrow(penguins_data), 0.5, 1.5)
  penguins_data$w <- penguins_data$w / mean(penguins_data$w)

  result <- compose_table(penguins_data, var = "species", weight = "w")
  expect_equal(sum(result$n, na.rm = TRUE), sum(penguins_data$w), tolerance = 1e-4)
})

# 4. na.rm behavior -------------------------------------------------------------

test_that("compose_table with na.rm = TRUE excludes NA observations from n sum", {
  penguins_with_na <- datasets::penguins  # has NAs in sex
  n_non_na <- sum(!is.na(penguins_with_na$sex))

  result <- suppressWarnings(
    compose_table(penguins_with_na, var = "sex", na.rm = TRUE)
  )
  expect_equal(sum(result$n, na.rm = TRUE), n_non_na)
})
