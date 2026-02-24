#' Compute adjusted standardized residuals for contingency tables
#'
#' Calculates adjusted standardized residuals for a contingency table formed
#' by crossing a row variable with one or more grouping (column) variables.
#' These residuals measure how much each observed cell frequency deviates from
#' what would be expected under independence, adjusted for both row and column
#' proportions.
#'
#' @param data A data frame or tibble containing the survey data.
#' @param var Character string specifying the row variable name.
#' @param group Character vector specifying one or more column grouping variable
#'   names. When multiple groups are provided, residuals are computed separately
#'   for each group's contingency table.
#' @param weight Character string specifying the weight variable name.
#'
#' @return A data.table with the same structure as the output from
#'   \code{compose_table()} with the same grouping variables. The first column
#'   contains the row variable levels, and subsequent columns contain adjusted
#'   standardized residuals for each group level. When multiple groups are
#'   specified, columns are prefixed with \code{[group_name]} to match
#'   \code{compose_table()} output.
#'
#' @details
#' Adjusted standardized residuals are calculated as:
#' \deqn{\frac{O - E}{\sqrt{E \times (1 - p_{row}) \times (1 - p_{col})}}}
#' where \eqn{O} is the observed frequency, \eqn{E} is the expected frequency
#' under independence, \eqn{p_{row}} is the row proportion, and \eqn{p_{col}}
#' is the column proportion.
#'
#' These residuals can be interpreted like z-scores: values beyond ±2 suggest
#' significant deviation from independence at approximately the 0.05 level.
#' They are equivalent to the standardized residuals from \code{chisq.test()}.
#'
#' When multiple grouping variables are provided, residuals are computed
#' independently for each two-way contingency table (var × group_i), then
#' combined into a single output table matching the structure of
#' \code{compose_table()}.
#'
#' @examples
#' \dontrun{
#' # Single group
#' compute_adj_residuals(cvvm, var = "t_VZD", group = "IDE_8", weight = "soc_weight")
#'
#' # Multiple groups
#' compute_adj_residuals(
#'   cvvm,
#'   var = "t_VZD",
#'   group = c("t_VEK_6", "IDE_8"),
#'   weight = "soc_weight"
#' )
#' }
#'
#' @export

compute_adj_residuals <- function(data, var, group, weight) {
  # Handle multiple groups by computing residuals for each separately

  if (length(group) > 1) {
    # Compute residuals for each group
    resid_list <- lapply(group, function(grp) {
      compute_adj_residuals_single(data, var, grp, weight)
    })

    # Combine: start with var column from first result
    result <- resid_list[[1]][, .(var)]

    # Add residual columns from each group
    for (i in seq_along(group)) {
      grp <- group[i]
      resid_cols <- grep("^\\[.+\\]", names(resid_list[[i]]), value = TRUE)
      result <- cbind(result, resid_list[[i]][, ..resid_cols])
    }

    return(result[])
  }

  # Single group case
  compute_adj_residuals_single(data, var, group, weight)
}

#' Compute adjusted residuals for a single grouping variable
#'
#' @param data A data frame or tibble containing the survey data.
#' @param var Character string specifying the row variable name.
#' @param group Character string specifying the column grouping variable name.
#' @param weight Character string specifying the weight variable name.
#'
#' @return A data.table with adjusted standardized residuals.
#' @keywords internal

compute_adj_residuals_single <- function(data, var, group, weight) {
  tab <- compose_table(data = data, var = var, group = group, weight = weight)
  tab <- tab[,
    .SD,
    .SDcols = c("var", grep("\\[.+\\]", names(tab), value = TRUE))
  ]

  # Extract the count matrix
  count_cols <- grep("\\[.+\\]", names(tab), value = TRUE)
  counts <- as.matrix(tab[, ..count_cols])

  # Calculate marginals
  row_totals <- rowSums(counts)
  col_totals <- colSums(counts)
  grand_total <- sum(counts)

  # Calculate expected frequencies
  expected <- outer(row_totals, col_totals) / grand_total

  # Calculate adjusted standardized residuals
  row_props <- row_totals / grand_total
  col_props <- col_totals / grand_total

  adj_resid <- (counts - expected) /
    sqrt(expected * outer(1 - row_props, 1 - col_props))

  # Convert back to data.table with same structure
  result <- data.table::copy(tab)
  result[, (count_cols) := data.table::as.data.table(adj_resid)]

  result[]
}
