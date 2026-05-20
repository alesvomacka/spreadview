#' Get names of categorical variables from a data frame
#'
#' Extracts the names of factor variables from a data frame. Optionally includes
#' character variables and allows exclusion of specific variable names.
#'
#' @param data A data frame to extract categorical variable names from.
#' @param exclude A character vector of variable names to exclude from the results.
#'   Default is `NULL` (no exclusions).
#' @param include_char Logical. If `TRUE`, character variables are included in
#'   addition to factor variables. Default is `FALSE`.
#'
#' @return A character vector of variable names.
#'
#' @examples
#' # Get only factor variables
#' get_categorical_vars(iris)
#'
#' # Include character variables
#' get_categorical_vars(mtcars, include_char = TRUE)
#'
#' # Exclude specific variables
#' get_categorical_vars(iris, exclude = c("Species"))
#'
#' @export
get_categorical_vars <- function(data, exclude = NULL, include_char = FALSE) {
  factor_vars <- names(data)[sapply(data, is.factor)]

  if (include_char) {
    char_vars <- names(data)[sapply(data, is.character)]
    factor_vars <- c(factor_vars, char_vars)
  }

  if (!is.null(exclude)) {
    factor_vars <- setdiff(factor_vars, exclude)
  }

  factor_vars
}
