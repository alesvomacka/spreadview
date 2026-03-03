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
#'   \item{item}{Variable name of the tabulated variable}
#'   \item{label}{Variable label from the `var` attribute (empty string if not available)}
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

  # Total sample size for this variable
  N_total <- sum(tab$n, na.rm = TRUE)

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

  # Add variable name as first column and label as second column
  var_label <- attr(data[[var]], "label")
  tab[, item := ""]
  tab[, label := ""]
  tab[1L, item := var]
  tab[1L, label := if (!is.null(var_label)) var_label else ""]
  data.table::setcolorder(tab, c("item", "label"))
  data.table::setnames(tab, old = var, new = "var")

  # Add N column (total sample size) before n — only in the first row
  tab[, N := NA_real_]
  tab[1L, N := N_total]
  data.table::setcolorder(tab, c("item", "label", "var", "N", "n"))

  # Add group totals row when grouping is used
  if (!is.null(group)) {
    group_col_names <- grep("^\\[.+\\]", names(tab), value = TRUE)
    totals_row <- data.table::data.table(
      item = NA_character_,
      label = NA_character_,
      var = NA_character_,
      n = NA_real_,
      N = NA_real_,
      total = NA_real_
    )
    # Compute absolute frequency for each group level
    for (gc in group_col_names) {
      # Extract group variable name and level from column name like "[island] Biscoe"
      grp_match <- regmatches(gc, regexec("^\\[(.+?)\\] (.+)$", gc))[[1]]
      if (length(grp_match) == 3) {
        grp_var_label <- grp_match[2]
        grp_level <- grp_match[3]
        # Find the actual group variable name (may differ from label)
        grp_var <- NULL
        for (g in group) {
          g_label <- attr(data[[g]], "label") %||% g
          if (g_label == grp_var_label) {
            grp_var <- g
            break
          }
        }
        if (!is.null(grp_var)) {
          if (is.null(weight)) {
            freq <- sum(data[[grp_var]] == grp_level, na.rm = TRUE)
          } else {
            freq <- sum(
              data[[weight]][data[[grp_var]] == grp_level],
              na.rm = TRUE
            )
          }
          data.table::set(totals_row, j = gc, value = freq)
        }
      }
    }
    tab <- data.table::rbindlist(list(tab, totals_row), fill = TRUE)
  }

  tab[]
}
