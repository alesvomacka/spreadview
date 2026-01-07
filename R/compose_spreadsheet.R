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
  weight,
  file = NULL,
  na.rm = TRUE
) {
  # Apply compose_table to each variable
  tables <- lapply(vars, function(var) {
    compose_table(
      data = data,
      var = var,
      group = group,
      weight = weight,
      prop = "col",
      na.rm = na.rm
    )
  })

  # Row bind all tables together
  spread <- data.table::rbindlist(tables, fill = TRUE)

  # Format "total" and all grouped columns as percentages
  total_idx <- which(names(spread) == "total")
  if (length(total_idx) > 0) {
    pct_cols <- names(spread)[total_idx:ncol(spread)]
    for (col in pct_cols) {
      class(spread[[col]]) <- c(class(spread[[col]]), "percentage")
    }
  }

  # Build header and create workbook
  wb <- openxlsx2::wb_workbook() |>
    openxlsx2::wb_add_worksheet() |>
    compose_headers(spread) |>
    # Add data starting at row 3
    openxlsx2::wb_add_data(x = spread, col_names = FALSE, start_row = 3)

  # Apply coloring based on residuals if grouping is specified
  if (!is.null(group)) {
    wb <- compose_residuals(wb, spread, data, vars, group, weight)
  }

  # Apply styling (borders, number formatting, text wrapping/centering)
  wb <- style_spreadsheet(wb, spread, group)

  if (!is.null(file)) {
    openxlsx2::wb_save(wb, file = file)
  }

  wb
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

style_spreadsheet <- function(wb, spread, group = NULL) {
  n_rows <- nrow(spread)
  n_cols <- ncol(spread)

  # Setting column width

  wb$set_col_widths(cols = 1, widths = 50)
  wb$set_col_widths(cols = 2, widths = 26)
  wb$set_col_widths(cols = 3:n_cols, widths = 10)

  # Format n column to show no decimals
  n_col <- which(names(spread) == "n")
  dims_n <- openxlsx2::wb_dims(rows = seq(3, n_rows + 2), cols = n_col)
  wb$add_numfmt(dims = dims_n, numfmt = "0")

  # Format percentage columns (from "total" onwards) to show no decimals
  total_col <- which(names(spread) == "total")
  if (length(total_col) > 0) {
    pct_cols <- seq(total_col, n_cols)
    dims_pct <- openxlsx2::wb_dims(
      rows = seq(3, n_rows + 2),
      cols = pct_cols
    )
    wb$add_numfmt(dims = dims_pct, numfmt = "0%")
  }

  # Wrap text in header cells
  dims_merged_header <- openxlsx2::wb_dims(rows = 1:2, cols = 1:n_cols)
  wb$add_cell_style(
    dims = dims_merged_header,
    wrap_text = TRUE,
    horizontal = "center",
    vertical = "top"
  )

  # Center numbers and percentages
  dims_values <- openxlsx2::wb_dims(rows = 3:(n_rows + 2), cols = 3:n_cols)
  wb$add_cell_style(dims = dims_values, horizontal = "center")

  # Identify columns that need left borders (for grouping variables)
  left_border_cols <- integer(0)
  if (!is.null(group)) {
    col_names <- names(spread)
    is_grouped <- grepl("^\\[.+\\] .+", col_names)

    # Extract group headers (e.g., "[Věková skupina]" from "[Věková skupina] 18-29 let")
    group_headers <- ifelse(
      is_grouped,
      sub("^(\\[.+\\]) .+$", "\\1", col_names),
      ""
    )

    # Find unique group headers and their first occurrence
    unique_groups <- unique(group_headers[is_grouped])
    for (grp_header in unique_groups) {
      first_col_idx <- which(group_headers == grp_header)[1]
      if (!is.na(first_col_idx)) {
        left_border_cols <- c(left_border_cols, first_col_idx)
        dims_border <- openxlsx2::wb_dims(
          rows = 1:(n_rows + 2),
          cols = first_col_idx
        )
        wb$add_border(
          dims = dims_border,
          left_border = "thick",
          left_color = openxlsx2::wb_color("000000"),
          right_border = NULL,
          top_border = NULL,
          bottom_border = NULL
        )
      }
    }
  }

  # Add thick right border to column 2
  dims_col2 <- openxlsx2::wb_dims(
    rows = 1:(n_rows + 2),
    cols = 2
  )
  wb$add_border(
    dims = dims_col2,
    right_border = "thick",
    right_color = openxlsx2::wb_color("000000")
  )

  # Add thick right border to last column
  dims_last_col <- openxlsx2::wb_dims(
    rows = 1:(n_rows + 2),
    cols = n_cols
  )
  # ...existing code for last column border...

  # Add top thick border to rows where label column is not empty

  # Apply AFTER other borders and handle special columns separately
  label_col <- spread[["label"]]
  label_rows <- which(!is.na(label_col) & label_col != "")

  # Columns that need special treatment (have left borders or right borders)
  special_cols <- c(2, left_border_cols)

  for (row_idx in label_rows) {
    excel_row <- row_idx + 2 # Account for 2 header rows

    # Apply top border to regular columns (not special)
    regular_cols <- setdiff(1:n_cols, special_cols)
    if (length(regular_cols) > 0) {
      dims_top_border <- openxlsx2::wb_dims(
        rows = excel_row,
        cols = regular_cols
      )
      wb$add_border(
        dims = dims_top_border,
        top_border = "thick",
        top_color = openxlsx2::wb_color("000000"),
        left_border = NULL,
        right_border = NULL,
        bottom_border = NULL
      )
    }

    # Apply top + left border to columns that have left borders
    for (col_idx in special_cols) {
      dims_special <- openxlsx2::wb_dims(rows = excel_row, cols = col_idx)
      wb$add_border(
        dims = dims_special,
        top_border = "thick",
        top_color = openxlsx2::wb_color("000000"),
        left_border = "thick",
        left_color = openxlsx2::wb_color("000000"),
        right_border = NULL,
        bottom_border = NULL
      )
    }

    # Apply top + right border to column 2
    dims_col2_top <- openxlsx2::wb_dims(rows = excel_row, cols = 2)
    wb$add_border(
      dims = dims_col2_top,
      top_border = "thick",
      top_color = openxlsx2::wb_color("000000"),
      right_border = "thick",
      right_color = openxlsx2::wb_color("000000"),
      left_border = NULL,
      bottom_border = NULL
    )
  }

  wb
}
