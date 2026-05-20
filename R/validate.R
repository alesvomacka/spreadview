#' Check Input Data and Variables
#'
#' Validates that the input data is a data frame and that specified variable
#' names exist in the data and are factors.
#'
#' @param data A data frame containing the variables to check.
#' @param vars A character vector of variable names that must exist in `data`
#'   and be factors.
#' @param group A character vector of grouping variable names that must exist
#'   in `data` and be factors.
#'
#' @return Invisibly returns `TRUE` if all checks pass.
#'
#' @keywords internal
check_input <- function(data, vars, group) {
  if (!is.data.frame(data)) {
    stop("'data' must be a data frame", call. = FALSE)
  }

  missing_vars <- setdiff(vars, names(data))
  if (length(missing_vars) > 0) {
    stop(
      "'vars' contains variables not found in data: ",
      paste(missing_vars, collapse = ", "),
      call. = FALSE
    )
  }

  non_factor_vars <- vars[!vapply(data[vars], is.factor, logical(1))]
  if (length(non_factor_vars) > 0) {
    stop(
      "'vars' contains non-factor variables: ",
      paste(non_factor_vars, collapse = ", "),
      call. = FALSE
    )
  }

  missing_groups <- setdiff(group, names(data))
  if (length(missing_groups) > 0) {
    stop(
      "'group' contains variables not found in data: ",
      paste(missing_groups, collapse = ", "),
      call. = FALSE
    )
  }

  non_factor_groups <- group[!vapply(data[group], is.factor, logical(1))]
  if (length(non_factor_groups) > 0) {
    stop(
      "'group' contains non-factor variables: ",
      paste(non_factor_groups, collapse = ", "),
      call. = FALSE
    )
  }

  invisible(TRUE)
}


#' Check Survey Weight Column
#'
#' Validates that a weight column in a dataset meets standard requirements for
#' survey weights. The function checks that weights are numeric, non-negative,
#' and contain no missing values. It also warns if the weights don't sum to the
#' number of observations or if the mean weight deviates from 1.
#'
#' @param data A data frame or data.table containing the weight column.
#' @param weight_col Character string specifying the name of the weight column
#'   to check.
#' @param tol Tolerance when checking weights sum and mean. Defaults to 1.
#'
#' @return Invisible NULL. The function is called for its side effects (errors
#'   and warnings).
#'
#' @keywords internal
check_weights <- function(data, weight_col, tol = 1) {
  if (!weight_col %in% names(data)) {
    stop("'weight_col' must be a column name in 'data'.")
  }

  weights <- data[[weight_col]]

  if (!is.numeric(weights)) {
    stop("'", weight_col, "' must be numeric.")
  }

  if (anyNA(weights)) {
    stop("'", weight_col, "' contains missing values.")
  }

  if (any(weights < 0)) {
    stop("'", weight_col, "' must contain non-negative values.")
  }

  data_nrow <- nrow(data)
  weights_sum <- sum(weights)
  weight_mean <- mean(weights)

  if (abs(data_nrow - weights_sum) > tol) {
    warning(
      "Weights sum (",
      weights_sum,
      ") is not equal to the number of observations (",
      data_nrow,
      ")."
    )
  }

  if (abs(1 - weight_mean) > tol) {
    warning("Weights mean should be 1, but is ", weight_mean, ".")
  }
}
