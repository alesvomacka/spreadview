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

  # Build workbook
  wb <- openxlsx2::wb_workbook() |>
    openxlsx2::wb_add_worksheet() |>
    compose_headers(spread) |>
    openxlsx2::wb_add_data(
      x = spread,
      col_names = FALSE,
      start_row = 3,
      na.strings = ""
    )

  # Apply residual coloring if grouped

  if (!is.null(group)) {
    wb <- compose_residuals(wb, spread)
  }

  wb <- style_spreadsheet(wb, spread, group, percent = percent)

  # Freeze headers
  wb$freeze_pane(first_active_row = 3, first_active_col = 7)

  if (!is.null(file)) {
    openxlsx2::wb_save(wb, file)
  }

  wb
}


#' Mark columns as percentage class
#'
#' @param spread A data.table with frequency data
#' @return The modified data.table with percentage class on relevant columns
#' @keywords internal
mark_percentage_cols <- function(spread) {
  total_idx <- which(names(spread) == "total")
  if (length(total_idx) == 0) {
    return(spread)
  }

  pct_cols <- names(spread)[total_idx:ncol(spread)]
  for (col in pct_cols) {
    class(spread[[col]]) <- c(class(spread[[col]]), "percentage")
  }
  spread
}


#' Style a Spreadsheet with Borders, Number Formatting, and Text Alignment
#'
#' Applies consistent styling to a workbook including column widths, number
#' formatting, text wrapping, centering, and borders.
#'
#' @param wb An \code{openxlsx2} workbook object.
#' @param spread A data.table containing the spreadsheet data (used to determine
#'   dimensions and column positions).
#' @param group Character vector specifying grouping variables. Default is
#'   \code{NULL}. Used to add borders between group sections.
#'
#' @return The modified workbook object with styling applied.
#'
#' @details
#' The function applies the following styling:
#' \itemize{
#'   \item Column widths: 50 for column 1, 26 for column 2, 10 for remaining
#'   \item Number formatting: no decimals for "n" column, percentage format for
#'     columns from "total" onwards
#'   \item Text wrapping and centering in header rows (1-2)
#'   \item Horizontal centering for value cells
#'   \item Thick left border on first column of each group variable
#'   \item Thick right border on column 2 and the last column
#' }
#'
#' @keywords internal

style_spreadsheet <- function(wb, spread, group = NULL, percent = TRUE) {
  n_rows <- nrow(spread)
  n_cols <- ncol(spread)

  # Setting column width

  wb$set_col_widths(cols = 1, widths = 10)
  wb$set_col_widths(cols = 2, widths = 26)
  wb$set_col_widths(cols = 3, widths = 26)
  wb$set_col_widths(cols = 4:n_cols, widths = 10)

  # Format n and N columns to show no decimals
  n_col <- which(names(spread) == "n")
  dims_n <- openxlsx2::wb_dims(rows = seq(3, n_rows + 2), cols = n_col)
  wb$add_numfmt(dims = dims_n, numfmt = "0")

  N_col <- which(names(spread) == "N")
  if (length(N_col) > 0) {
    dims_N <- openxlsx2::wb_dims(rows = seq(3, n_rows + 2), cols = N_col)
    wb$add_numfmt(dims = dims_N, numfmt = "0")
  }

  # Format percentage columns (from "total" onwards) to show no decimals
  total_col <- which(names(spread) == "total")
  if (length(total_col) > 0) {
    pct_cols <- seq(total_col, n_cols)
    dims_pct <- openxlsx2::wb_dims(
      rows = seq(3, n_rows + 2),
      cols = pct_cols
    )
    wb$add_numfmt(dims = dims_pct, numfmt = if (percent) "0%" else "0")
  }

  # Override formatting for group totals rows (absolute counts, not percentages)
  if (!is.null(group)) {
    totals_rows <- which(is.na(spread[["item"]]) & is.na(spread[["var"]]))
    if (length(totals_rows) > 0) {
      group_col_indices <- grep("^\\[.+\\]", names(spread))
      if (length(group_col_indices) > 0) {
        dims_totals <- openxlsx2::wb_dims(
          rows = totals_rows + 2,
          cols = group_col_indices
        )
        wb$add_numfmt(dims = dims_totals, numfmt = "0")
      }
    }
  }

  # Wrap text in header cells
  dims_merged_header <- openxlsx2::wb_dims(rows = 1:2, cols = 1:n_cols)
  wb$add_cell_style(
    dims = dims_merged_header,
    wrap_text = TRUE,
    horizontal = "center",
    vertical = "top"
  )

  # Center numbers and percentages (cols 4 onwards: skip item, label, var)
  dims_values <- openxlsx2::wb_dims(rows = 3:(n_rows + 2), cols = 4:n_cols)
  wb$add_cell_style(dims = dims_values, horizontal = "center")

  # Identify columns that need left borders (for grouping variables)
  left_border_cols <- integer(0)
  if (!is.null(group)) {
    col_names <- names(spread)
    is_grouped <- grepl("^\\[.+\\] .+", col_names)
    group_headers <- ifelse(
      is_grouped,
      sub("^(\\[.+\\]) .+$", "\\1", col_names),
      ""
    )
    unique_groups <- unique(group_headers[is_grouped])
    for (grp_header in unique_groups) {
      first_col_idx <- which(group_headers == grp_header)[1]
      if (!is.na(first_col_idx)) {
        left_border_cols <- c(left_border_cols, first_col_idx)
      }
    }
  }

  right_border_cols <- intersect(c(3L, n_cols), seq_len(n_cols))
  plain_cols        <- setdiff(seq_len(n_cols), c(right_border_cols, left_border_cols))
  left_only_cols    <- setdiff(left_border_cols, right_border_cols)
  right_has_left    <- vapply(right_border_cols, `%in%`, logical(1L), left_border_cols)

  item_col   <- spread[["item"]]
  label_rows <- which(!is.na(item_col) & item_col != "")
  all_rows   <- 1:(n_rows + 2)

  black <- openxlsx2::wb_color("000000")

  # --- Permanent column borders (full contiguous range, one call each) ---
  for (col in left_border_cols) {
    wb$add_border(
      dims          = openxlsx2::wb_dims(rows = all_rows, cols = col),
      left_border   = "thick", left_color = black,
      right_border  = NULL, top_border = NULL, bottom_border = NULL
    )
  }
  wb$add_border(
    dims          = openxlsx2::wb_dims(rows = all_rows, cols = 3),
    right_border  = "thick", right_color = black
  )
  wb$add_border(
    dims          = openxlsx2::wb_dims(rows = all_rows, cols = n_cols),
    right_border  = "thick", right_color = black
  )

  # --- Thick top border on the first row of each variable block ---
  for (row_idx in label_rows) {
    excel_row <- row_idx + 2L

    if (length(plain_cols) > 0) {
      wb$add_border(
        dims          = openxlsx2::wb_dims(rows = excel_row, cols = plain_cols),
        top_border    = "thick", top_color = black,
        left_border   = NULL, right_border = NULL, bottom_border = NULL
      )
    }

    for (i in seq_along(right_border_cols)) {
      wb$add_border(
        dims          = openxlsx2::wb_dims(rows = excel_row, cols = right_border_cols[i]),
        top_border    = "thick", top_color = black,
        right_border  = "thick", right_color = black,
        left_border   = if (right_has_left[i]) "thick" else NULL,
        left_color    = if (right_has_left[i]) black else NULL,
        bottom_border = NULL
      )
    }

    for (col in left_only_cols) {
      wb$add_border(
        dims          = openxlsx2::wb_dims(rows = excel_row, cols = col),
        top_border    = "thick", top_color = black,
        left_border   = "thick", left_color = black,
        right_border  = NULL, bottom_border = NULL
      )
    }
  }

  wb
}
