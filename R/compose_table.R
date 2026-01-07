#' Compose a frequency table with optional grouping
#'
#' Creates a frequency table for a categorical variable, optionally cross-tabulated
#' by one or more grouping variables. The function automatically adds variable labels
#' and standardizes column names for easier downstream processing.
#'
#' @param data A data frame or data.table containing the variables to tabulate.
#' @param var Character string specifying the name of the variable to tabulate.
#' @param group Character vector of variable name(s) to group by. If `NULL` (default),
#'   creates a simple frequency table. If provided, creates cross-tabulations with
#'   each grouping variable. Group columns are prefixed with `[group_name]` in the
#'   output.
#' @param weight Character string specifying the name of the weight variable to use
#'   for weighted frequencies.
#' @param prop Character string specifying proportion type. Options are `"none"` for
#'   counts only, `"col"` for column proportions, or `"row"` for row proportions.
#'   Default is `"none"`.
#'
#' @return A data.table with the following structure:
#'   \item{label}{Variable label from the `var` attribute (if available)}
#'   \item{var}{Factor levels of the tabulated variable}
#'   \item{n}{Weighted counts}
#'   \item{total}{Overall proportions (if `prop != "none"`)}
#'   \item{[group] level}{One column per level of each grouping variable (if `group` is specified)}
#'
#' @examples
#' # Simple frequency table
#' compose_table(cvvm, var = "t_VZD", weight = "soc_weight", prop = "col")
#'
#' # Cross-tabulation with one grouping variable
#' compose_table(cvvm, var = "t_VZD", weight = "soc_weight",
#'               prop = "col", group = "t_VEK_6")
#'
#' # Cross-tabulation with multiple grouping variables
#' compose_table(cvvm, var = "t_VZD", weight = "soc_weight",
#'               prop = "col", group = c("t_VEK_6", "IDE_8"))
#'
#' @export
#'

compose_table <- function(
  data,
  var,
  group = NULL,
  weight,
  prop = "none",
  na.rm = TRUE
) {
  if (is.null(group)) {
    tab <- count_freqs(
      data = data,
      var = var,
      weight = weight,
      prop = prop,
      na.rm = na.rm
    )
  } else {
    # Start with total
    tab <- count_freqs(
      data = data,
      var = var,
      weight = weight,
      prop = prop
    )

    # Join each group variable
    for (grp in group) {
      tab_grouped <- count_freqs(
        data = data,
        var = var,
        weight = weight,
        prop = prop,
        group = grp
      )

      # Get group columns (excluding var and n)
      group_cols <- setdiff(names(tab_grouped), c(var, "n"))

      # Get group variable label (if available) or use variable name
      grp_label <- attr(data[[grp]], "label")
      grp_prefix <- if (!is.null(grp_label)) grp_label else grp

      # Rename group columns to include group variable name/label
      for (col in group_cols) {
        data.table::setnames(
          tab_grouped,
          old = col,
          new = paste0("[", grp_prefix, "] ", col)
        )
      }

      # Join to main table
      tab <- tab[tab_grouped, on = var, env = list(var = var)]
    }
  }

  # Add variable label
  var_label <- attr(data[[var]], "label")
  tab$label <- ""
  tab$label[1] <- if (!is.null(var_label)) var_label else var
  data.table::setcolorder(tab, "label")

  # Standardize var column name
  data.table::setnames(tab, old = var, new = "var")

  tab[]
}
