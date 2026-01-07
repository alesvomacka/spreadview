#'  Calculate Frequency Tables with Optional Proportions
#'
#' Computes weighted frequency tables for a variable, optionally grouped by
#' another variable. Can return counts or proportions calculated across
#' different margins (total, row, or column).
#'
#' @param data A data frame or object that can be coerced to a data.table.
#' @param var Character string specifying the name of the variable to tabulate.
#' @param group Character string specifying the name of the grouping variable,
#'   or `NULL` (default) for no grouping.
#' @param weight Character string specifying the name of the weight variable
#'   to use for weighted frequencies. If no weights are specified, equal weights
#'   are assumed.
#' @param prop Character string indicating the type of proportion to calculate.
#'   Options are:
#'   * `"none"` (default): Return weighted counts only
#'   * `"total"`: Proportions of the grand total
#'   * `"row"`: Proportions within each row (within each level of `var`)
#'   * `"col"`: Proportions within each column (within each level of `group`)
#' @param na.rm Logical. If `TRUE`, removes rows with missing values in `var`
#'   before computing frequencies. Default is `FALSE`.
#'
#' @return A data.table containing:
#'   * When `group` is `NULL`: A table with `var` and either `n` (counts) or
#'     `total` (proportions)
#'   * When `group` is specified and `prop != "none"`: A wide-format table with
#'     `var` as rows and `group` levels as columns, containing proportions
#'   * When `group` is specified and `prop == "none"`: A long-format table with
#'     `var`, `group`, and `n` columns
#'
#' @examples
#' count_freqs(data = penguins, var = "species")
#' count_freqs(data = penguins, var = "species", group = "island", prop = "")
#'
#' @export
#' @importFrom data.table .SD .N

count_freqs <- function(
  data,
  var,
  group = NULL,
  weight = NULL,
  prop = "none",
  na.rm = TRUE
) {
  if (is.null(weight)) {
    data[[weight]] <- 1
    warning("No weight specified. Assuming equal weights for all observations.")
  }

  check_weights(data, weight_col = weight)
  if (!prop %in% c("none", "total", "row", "col")) {
    stop('prop has to be either "none", "total", "row" "col"')
  }

  data <- data.table::data.table(data)

  if (na.rm) {
    data <- data[!is.na(data[[var]])]
  }

  if (is.null(group)) {
    # No grouping var
    tab <- data[,
      .(n = sum(weight)),
      by = var,
      env = list(var = var, weight = weight, group = group)
    ]

    if (prop %in% c("total", "col", "row")) {
      tab[, total := n / sum(n)]
    }
  } else {
    # With grouping var
    tab <- data[,
      .(n = sum(weight)),
      by = list(var, group),
      env = list(var = var, weight = weight, group = group)
    ]

    if (prop == "total") {
      tab[, prop := n / sum(n)]
      form <- as.formula(paste(var, "~", group))
      tab <- data.table::dcast(tab, form, value.var = "prop", fill = 0)
    } else if (prop == "row") {
      tab[,
        prop := n / sum(n),
        by = var,
        env = list(var = var, weight = weight, group = group)
      ]
      form <- as.formula(paste(var, "~", group))
      tab <- data.table::dcast(tab, form, value.var = "prop", fill = 0)
    } else if (prop == "col") {
      tab[,
        prop := n / sum(n),
        by = group,
        env = list(var = var, weight = weight, group = group)
      ]
      form <- as.formula(paste(var, "~", group))
      tab <- data.table::dcast(tab, form, value.var = "prop", fill = 0)
    } else if (prop == "none") {
      form <- as.formula(paste(var, "~", group))
      tab <- data.table::dcast(tab, form, value.var = "n", fill = 0)
    }
  }

  tab[]
}
