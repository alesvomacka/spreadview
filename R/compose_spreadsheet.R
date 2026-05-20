#' Compose a Spreadsheet with Multiple Frequency Tables
#'
#' Creates an Excel workbook containing frequency tables for multiple variables,
#' optionally grouped and weighted. Each variable's table is row-bound together
#' into a single sheet. When grouping is specified, cells are colored based on
#' adjusted standardized residuals.
#'
#' @param data A data frame or data.table containing the variables to tabulate.
#' @param vars Character vector of variable names to create tables for. Each
#'   variable will be passed to \code{compose_table()}.
#' @param group Character vector specifying grouping variables. Default is
#'   \code{NULL} for no grouping. Passed to \code{compose_table()}.
#' @param weight Character string naming the weight variable to use for
#'   weighted frequencies. Passed to \code{compose_table()}.
#' @param file Character string specifying the output file path. If \code{NULL}
#'   (default), returns the workbook object without saving.
#' @param na.rm Should missing values be removed from the table? Defaults to TRUE.
#' @param percent Logical. If \code{TRUE} (default), percentage values are
#'   formatted as percentages (e.g., 50\%). If \code{FALSE}, they are formatted
#'   as plain numbers multiplied by 100 (e.g., 50).
#'
#' @return An \code{openxlsx2} workbook object containing a single worksheet
#'   with all frequency tables row-bound together. Cells are colored based on
#'   adjusted standardized residuals when grouping is used.
#'
#' @details
#' The function iterates over each variable in \code{vars}, applies
#' \code{compose_table()} with the specified parameters, and combines the
#' results using \code{data.table::rbindlist()}. The resulting spreadsheet
#' is added to a new workbook with \code{openxlsx2}.
#'
#' The header spans two rows:
#' \itemize{
#'   \item Row 1: Column names for non-group columns (e.g., "label", "n", "total")
#'     and group variable names for grouped columns (e.g., "IDE_8")
#'   \item Row 2: Empty strings for non-group columns, category names for
#'     grouped columns (e.g., "muž", "žena")
#' }
#'
#' When grouping variables are specified, cells are colored based on adjusted
#' standardized residuals:
#' \itemize{
#'   \item Blue shades for positive residuals (more than expected)
#'   \item Red shades for negative residuals (less than expected)
#'   \item Color intensity increases with absolute residual value
#'   \item Residuals beyond ±2 indicate significance at ~0.05 level
#' }
#'
#' @examples
#' \dontrun{
#' compose_spreadsheet(data = penguins, var = c("species", "sex"), group = "island")
#' }
#'
#' @seealso \code{\link{compose_table}}, \code{\link{compute_adj_residuals}}
#'
#' @export

compose_spreadsheet <- function(
  data,
  vars,
  group = NULL,
  weight = NULL,
  file = NULL,
  na.rm = TRUE,
  percent = TRUE
) {
  # Check inputs
  check_input(data = data, vars = vars, group = group)

  # Build all tables and combine
  spread <- vars |>
    lapply(\(var) {
      compose_table(data, var, group, weight, prop = "col", na.rm = na.rm)
    }) |>
    data.table::rbindlist(fill = TRUE)

  # Mark percentage columns
  spread <- mark_percentage_cols(spread)

  # Build workbook skeleton (data not yet added)
  wb <- openxlsx2::wb_workbook() |>
    openxlsx2::wb_add_worksheet() |>
    compose_headers(spread)

  # Apply residual coloring before the percent transformation so that
  # compose_residuals always sees proportions in the 0–1 range, not 0–100.
  if (!is.null(group)) {
    wb <- compose_residuals(wb, spread)
  }

  # When not using percent format, multiply percentage values by 100
  if (!percent) {
    pct_cols <- names(spread)[vapply(
      spread,
      inherits,
      logical(1),
      "percentage"
    )]
    # Group total rows have total = NA (absolute counts) — skip them
    pct_rows <- which(!is.na(spread[["total"]]))
    for (col in pct_cols) {
      data.table::set(spread, pct_rows, col, spread[[col]][pct_rows] * 100)
    }
  }

  wb <- openxlsx2::wb_add_data(
    wb,
    x = spread,
    col_names = FALSE,
    start_row = 3,
    na.strings = ""
  )

  wb <- style_spreadsheet(wb, spread, group, percent = percent)

  # Freeze headers
  wb$freeze_pane(first_active_row = 3, first_active_col = 7)

  if (!is.null(file)) {
    openxlsx2::wb_save(wb, file)
  }

  wb
}
