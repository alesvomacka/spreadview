# Internal helpers for adjusted standardized residual calculations.

# Compute adjusted residuals for a single grouping variable.
# Called by compute_adj_residuals() for each group in turn.
#
# @keywords internal
compute_adj_residuals_single <- function(data, var, group, weight) {
  tab <- compose_table(data = data, var = var, group = group, weight = weight)
  tab <- tab[,
    .SD,
    .SDcols = c("var", grep("\\[.+\\]", names(tab), value = TRUE))
  ]

  tab_nototal <- tab <- tab[!nrow(tab), ] # drop the last row, which holds column totals

  count_cols <- grep("\\[.+\\]", names(tab_nototal), value = TRUE)
  counts <- as.matrix(tab_nototal[, ..count_cols])

  row_totals <- rowSums(counts)
  col_totals <- colSums(counts)
  grand_total <- sum(counts)

  expected <- outer(row_totals, col_totals) / grand_total

  row_props <- row_totals / grand_total
  col_props <- col_totals / grand_total

  adj_resid <- (counts - expected) /
    sqrt(expected * outer(1 - row_props, 1 - col_props))

  result <- data.table::copy(tab)
  result[, (count_cols) := data.table::as.data.table(adj_resid)]

  result[]
}


# Map standardized residual values to fill colours.
# Returns NA_character_ where no colour applies.
residual_to_color_vec <- function(resid) {
  result <- rep(NA_character_, length(resid))
  finite <- !is.na(resid) & is.finite(resid)
  result[finite & resid >= 2.3]                  <- "#6BAED6"
  result[finite & resid >= 1.9 & resid < 2.3]   <- "#BDD7E7"
  result[finite & resid <= -2.3]                 <- "#FB6A4A"
  result[finite & resid <= -1.9 & resid > -2.3] <- "#FCBBA1"
  result
}
