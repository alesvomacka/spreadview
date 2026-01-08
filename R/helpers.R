#' Extract pattern from character vector
#'
#' Extracts substrings matching a regular expression pattern from a character
#' vector. For elements that don't match the pattern, the original value is
#' returned unchanged.
#'
#' @param x A character vector to extract patterns from.
#' @param pattern A regular expression pattern to match and extract.
#'
#' @return A character vector of the same length as `x`, where matching elements
#'   contain the extracted pattern and non-matching elements contain their
#'   original values.
#'
#' @examples
#' # Extract text within square brackets
#' x <- c("[IDE_8] muž", "[IDE_8] žena", "no brackets")
#' extract_pattern(x, pattern = "\\[.+?\\]")
#' # Returns: c("[IDE_8]", "[IDE_8]", "no brackets")
#'
#' @export

extract_pattern <- function(x, pattern) {
  if (!is.character(x)) {
    stop("x has to be character vector")
  }

  ifelse(
    grepl(pattern, x),
    yes = regmatches(x, regexpr(pattern, x)),
    no = x
  )
}
extract_pattern <- function(x, pattern) {
  if (!is.character(x)) {
    stop("x has to be character vector")
  }

  ifelse(
    grepl(pattern, x),
    yes = regmatches(x, regexpr(pattern, x)),
    no = x
  )
}

#' Extract text outside a pattern from character vector
#'
#' Removes substrings matching a regular expression pattern (and any trailing
#' whitespace) from a character vector. For elements that don't match the
#' pattern, an empty string is returned.
#'
#' @param x A character vector to process.
#' @param pattern A regular expression pattern to match and remove.
#'
#' @return A character vector of the same length as `x`, where matching elements
#'   contain the remaining text after pattern removal, and non-matching elements
#'   contain empty strings.
#'
#' @examples
#' # Remove text within square brackets
#' x <- c("[IDE_8] muž", "[IDE_8] žena", "no brackets")
#' extract_pattern_outside(x, pattern = "\\[.+?\\]")
#' # Returns: c("muž", "žena", "")
#'
#' @export

extract_pattern_outside <- function(x, pattern) {
  if (!is.character(x)) {
    stop("x has to be character vector")
  }

  result <- sub(pattern = paste0(pattern, "\\s*"), replacement = "", x)
  ifelse(result == x, "", result)
}


#' Apply residual-based coloring to grouped columns
#'
#' Colors cells in grouped columns based on adjusted standardized residuals.
#' Blue shades indicate positive residuals (more than expected), red shades
#' indicate negative residuals (less than expected).
#'
#' @param wb An openxlsx2 workbook object with data already added.
#' @param spread A data.table containing the frequency table data.
#' @param data The original data frame used to compute residuals.
#' @param vars Character vector of variable names that were tabulated.
#' @param group Character vector of grouping variables.
#' @param weight Character string naming the weight variable.
#'
#' @return The modified workbook object with cell coloring applied.
#'
#' @keywords internal
compose_residuals <- function(wb, spread, data, vars, group, weight) {
  # Find group columns in the spread
  group_cols <- grep("^\\[.+\\]", names(spread), value = TRUE)

  if (length(group_cols) == 0) {
    return(wb)
  }

  # Track current row position (start at 3 for two header rows)
  current_row <- 3

  for (var in vars) {
    # Compute residuals for this variable (handles multiple groups)
    resid <- compute_adj_residuals(
      data = data,
      var = var,
      group = group,
      weight = weight
    )

    n_rows <- nrow(resid)
    resid_cols <- grep("^\\[.+\\]", names(resid), value = TRUE)

    # Apply coloring for each cell
    for (i in seq_len(n_rows)) {
      for (j in seq_along(resid_cols)) {
        resid_col_name <- resid_cols[j]
        resid_val <- resid[[resid_col_name]][i]

        if (!is.na(resid_val) && is.finite(resid_val)) {
          # Determine color based on residual value
          fill_color <- residual_to_color(resid_val)

          if (!is.null(fill_color)) {
            excel_row <- current_row + i - 1
            # Match residual column to spread column by name
            excel_col <- which(names(spread) == resid_col_name)

            if (length(excel_col) == 1) {
              wb <- wb |>
                openxlsx2::wb_add_fill(
                  dims = openxlsx2::wb_dims(
                    rows = excel_row,
                    cols = excel_col
                  ),
                  color = openxlsx2::wb_color(fill_color)
                )
            }
          }
        }
      }
    }

    current_row <- current_row + n_rows
  }

  wb
}

#' Convert residual value to fill color
#'
#' @param resid Numeric residual value
#' @return Hex color string or NULL for no coloring
#' @keywords internal
residual_to_color <- function(resid) {
  if (is.na(resid) || !is.finite(resid)) {
    return(NULL)
  }

  # Define color thresholds and corresponding colors
  # Positive residuals: blue shades (more than expected)
  # Negative residuals: red shades (less than expected)
  if (resid >= 2.3) {
    "#6BAED6"
  } else if (resid >= 1.9) {
    "#BDD7E7"
  } else if (resid <= -1.9) {
    "#FCBBA1"
  } else if (resid <= -2.3) {
    "#FB6A4A"
  } else {
    NULL
  }
}


#' Compose two-row headers for grouped spreadsheet
#'
#' Creates a two-row header structure where the first row contains column names
#' for non-group columns and group variable names for grouped columns, and the
#' second row contains category names for grouped columns.
#'
#' @param wb An openxlsx2 workbook object with at least one worksheet.
#' @param spread A data.table containing the frequency table data.
#'
#' @return The modified workbook object with headers added and merged cells.
#'
#' @keywords internal
compose_headers <- function(wb, spread) {
  col_names <- names(spread)
  is_grouped <- grepl("^\\[.+\\] .+", col_names)

  header_row1 <- ifelse(
    is_grouped,
    sub("^(\\[.+\\]) .+$", "\\1", col_names),
    col_names
  )
  header_row2 <- ifelse(
    is_grouped,
    sub("^\\[.+\\] (.+)$", "\\1", col_names),
    ""
  )

  # Create header data frame
  header_df <- data.frame(
    rbind(header_row1, header_row2),
    stringsAsFactors = FALSE
  )
  names(header_df) <- col_names

  # Add header rows (without column names)
  wb <- wb |>
    openxlsx2::wb_add_data(x = header_df, col_names = FALSE, start_row = 1)

  # Merge adjacent cells with the same group variable name in header row 1
  if (any(is_grouped)) {
    # Find runs of identical group names
    run_start <- NULL
    run_value <- NULL

    for (i in seq_along(header_row1)) {
      current_val <- header_row1[i]
      is_group_col <- is_grouped[i]

      if (is_group_col) {
        if (is.null(run_start) || current_val != run_value) {
          # End previous run if it spans multiple columns
          if (!is.null(run_start) && (i - 1) > run_start) {
            wb <- wb |>
              openxlsx2::wb_merge_cells(
                dims = openxlsx2::wb_dims(rows = 1, cols = run_start:(i - 1))
              )
          }
          # Start new run
          run_start <- i
          run_value <- current_val
        }
      } else {
        # End previous run if it exists and spans multiple columns
        if (
          !is.null(run_start) && (i - 1) >= run_start && (i - 1) > run_start
        ) {
          wb <- wb |>
            openxlsx2::wb_merge_cells(
              dims = openxlsx2::wb_dims(rows = 1, cols = run_start:(i - 1))
            )
        }
        run_start <- NULL
        run_value <- NULL
      }
    }

    # Handle final run at end of header
    if (!is.null(run_start) && length(header_row1) > run_start) {
      wb <- wb |>
        openxlsx2::wb_merge_cells(
          dims = openxlsx2::wb_dims(
            rows = 1,
            cols = run_start:length(header_row1)
          )
        )
    }
  }

  wb
}


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
#' @details
#' The function performs the following checks:
#' \enumerate{
#'   \item `data` is a data frame
#'   \item All variables in `vars` exist in `data`
#'   \item All variables in `vars` are factors
#'   \item All variables in `group` exist in `data`
#'   \item All variables in `group` are factors
#' }
#'
#' If any check fails, the function stops with an informative error message.
#'
#' @examples
#' \dontrun{
#' df <- data.frame(
#'   x = factor(c("a", "b", "c")),
#'   y = factor(c("d", "e", "f")),
#'   z = factor(c("g", "h", "i"))
#' )
#' check_input(df, vars = c("x", "y"), group = "z")
#' }
#'
#' @export

check_input <- function(data, vars, group) {
  # Check data is a data frame
  if (!is.data.frame(data)) {
    stop("'data' must be a data frame", call. = FALSE)
  }

  # Check vars are names of variables in data
  missing_vars <- setdiff(vars, names(data))
  if (length(missing_vars) > 0) {
    stop(
      "'vars' contains variables not found in data: ",
      paste(missing_vars, collapse = ", "),
      call. = FALSE
    )
  }

  # Check var variables are all factors
  non_factor_vars <- vars[!vapply(data[vars], is.factor, logical(1))]
  if (length(non_factor_vars) > 0) {
    stop(
      "'vars' contains non-factor variables: ",
      paste(non_factor_vars, collapse = ", "),
      call. = FALSE
    )
  }

  # Check group are names of variables in data
  missing_groups <- setdiff(group, names(data))
  if (length(missing_groups) > 0) {
    stop(
      "'group' contains variables not found in data: ",
      paste(missing_groups, collapse = ", "),
      call. = FALSE
    )
  }

  # Check group variables are all factors
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
