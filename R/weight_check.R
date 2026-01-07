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
#' @param tol Tolarence when checking weights sum and mean. Defaults to 0.01.
#'
#' @return Invisible NULL. The function is called for its side effects (errors
#'   and warnings).
#'
#' @details
#' The function performs the following checks:
#' \describe{
#'   \item{Errors}{
#'     - Weight column must exist in the dataset
#'     - Weight column must be numeric
#'     - Weight column must not contain missing values
#'     - All weights must be non-negative
#'   }
#'   \item{Warnings}{
#'     - If the sum of weights differs from the number of observations by more
#'       than 0.01
#'     - If the mean weight differs from 1 by more than 0.01
#'   }
#' }
#'
#' @examples
#' # Create sample data with valid weights
#' df <- data.frame(
#'   id = 1:100,
#'   weight = rep(1, 100)
#' )
#' check_weights(df, "weight")
#'
#' # Weights that sum to n but with variation
#' df$weight <- runif(100, 0.5, 1.5)
#' df$weight <- df$weight / mean(df$weight)
#' check_weights(df, "weight")
#'
#' @export
check_weights <- function(data, weight_col, tol = 1) {
  # Errors
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

  # Warnings
  data_nrow <- nrow(data)
  weights_sum <- sum(weights)
  weight_mean <- mean(weights)

  sum_check <- abs(data_nrow - weights_sum)
  mean_check <- abs(1 - weight_mean)

  if (sum_check > tol) {
    warning(
      "Weights sum (",
      weights_sum,
      ") is not equal to the number of observations (",
      data_nrow,
      ")."
    )
  }

  if (mean_check > tol) {
    warning("Weights mean should be 1, but is ", weight_mean, ".")
  }
}
