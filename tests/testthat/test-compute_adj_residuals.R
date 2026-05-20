library(testthat)

# Setup: use penguins without NAs for a clean contingency table
penguins_data <- na.omit(datasets::penguins)

# 1. Unweighted residuals match chisq.test()$stdres --------------------------

test_that("unweighted residuals match chisq.test()$stdres", {
  result <- suppressWarnings(compute_adj_residuals(
    penguins_data,
    var    = "species",
    group  = "island",
    weight = NULL
  ))

  ct <- chisq.test(table(penguins_data$species, penguins_data$island))
  expected_stdres <- ct$stdres

  # Pull residual columns from result (drop last totals row)
  resid_cols <- grep("^\\[.+\\]", names(result), value = TRUE)
  resid_matrix <- as.matrix(result[, ..resid_cols])

  # Column order from compute_adj_residuals may differ from table(); align by name
  island_levels <- colnames(expected_stdres)
  for (lvl in island_levels) {
    col_name <- resid_cols[grepl(paste0("\\] ", lvl, "$"), resid_cols)]
    expect_equal(
      as.numeric(resid_matrix[, col_name]),
      as.numeric(expected_stdres[, lvl]),
      tolerance = 1e-8,
      label = paste("residuals for island level", lvl)
    )
  }
})
