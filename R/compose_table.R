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
#' @param na.rm Should missing values be removed from the table? Defaults to TRUE.
#'
#' @return A data.table with the following structure:
#'   \item{label}{Variable label from the `var` attribute (if available)}
#'   \item{var}{Factor levels of the tabulated variable}
#'   \item{n}{Weighted counts}
#'   \item{total}{Overall proportions (if `prop != "none"`)}
#'   \item{[group] level}{One column per level of each grouping variable (if `group` is specified)}
#'
#' @examples
#' compose_table(penguins, var = "species", group = "island", prop = "col")
#' @export

compose_table <- function(
  data,
  var,
  group = NULL,
  weight = NULL,
  prop = "none",
  na.rm = TRUE
) {
  # Base table (total frequencies)
  tab <- count_freqs(
    data = data,
    var = var,
    weight = weight,
    prop = prop,
    na.rm = na.rm
  )

  # Add grouped cross-tabulations
  if (!is.null(group)) {
    for (grp in group) {
      tab_grouped <- count_freqs(
        data = data,
        var = var,
        weight = weight,
        prop = prop,
        group = grp,
        na.rm = na.rm
      )

      # Rename group columns with prefix
      group_cols <- setdiff(names(tab_grouped), c(var, "n"))
      grp_label <- attr(data[[grp]], "label") %||% grp
      new_names <- paste0("[", grp_label, "] ", group_cols)
      data.table::setnames(tab_grouped, old = group_cols, new = new_names)

      # Join to main table
      tab <- tab[tab_grouped, on = var]
    }
  }

  # Add variable label as first column
  var_label <- attr(data[[var]], "label") %||% var
  tab[, label := ""]
  tab[1L, label := var_label]
  data.table::setcolorder(tab, "label")
  data.table::setnames(tab, old = var, new = "var")

  tab[]
}
