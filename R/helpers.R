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


# Convert a 1-based column index to an Excel column letter (supports up to ZZ)
int_to_col <- function(n) {
  if (n <= 26L) return(LETTERS[n])
  paste0(LETTERS[(n - 1L) %/% 26L], LETTERS[(n - 1L) %% 26L + 1L])
}


# Vectorised version of residual-to-colour mapping.
# Returns NA_character_ where no colour applies.
residual_to_color_vec <- function(resid) {
  result <- rep(NA_character_, length(resid))
  finite <- !is.na(resid) & is.finite(resid)
  result[finite & resid >= 2.3]                    <- "#6BAED6"
  result[finite & resid >= 1.9 & resid < 2.3]     <- "#BDD7E7"
  result[finite & resid <= -2.3]                   <- "#FB6A4A"
  result[finite & resid <= -1.9 & resid > -2.3]   <- "#FCBBA1"
  result
}

# Apply wb_add_fill for each run of consecutive rows with the same colour in
# a single column. openxlsx2 requires rectangular dims ("A1" or "A1:A9"),
# so we encode runs to minimise the number of API calls.
apply_column_fills <- function(wb, col_letter, rows, colors) {
  n <- length(rows)
  if (n == 0L) return(wb)

  run_start <- 1L
  for (k in seq_len(n)) {
    end_run <- k == n ||
      colors[k + 1L] != colors[k] ||
      rows[k + 1L]   != rows[k] + 1L
    if (end_run) {
      r1 <- rows[run_start]
      r2 <- rows[k]
      dims_str <- if (r1 == r2) {
        paste0(col_letter, r1)
      } else {
        paste0(col_letter, r1, ":", col_letter, r2)
      }
      wb <- openxlsx2::wb_add_fill(
        wb,
        dims  = dims_str,
        color = openxlsx2::wb_color(colors[k])
      )
      run_start <- k + 1L
    }
  }
  wb
}


#' Apply residual-based coloring to grouped columns
#'
#' Colors cells in grouped columns based on adjusted standardized residuals.
#' Blue shades indicate positive residuals (more than expected), red shades
#' indicate negative residuals (less than expected).
#'
#' @param wb An openxlsx2 workbook object with data already added.
#' @param spread A data.table containing the frequency table data.
#'
#' @return The modified workbook object with cell coloring applied.
#'
#' @keywords internal
compose_residuals <- function(wb, spread) {
  group_cols <- grep("^\\[.+\\]", names(spread), value = TRUE)
  if (length(group_cols) == 0L) return(wb)

  # Determine unique group variable labels from column names like "[label] level"
  group_labels <- unique(sub("^\\[(.+?)\\] .+$", "\\1", group_cols))

  # Accumulate (excel_row, color) pairs per spread column across all variables,
  # then flush column-by-column using run-length encoding so consecutive rows
  # of the same colour become a single "A5:A9" range → one wb_add_fill call.
  col_rows   <- vector("list", length(group_cols))
  col_colors <- vector("list", length(group_cols))
  names(col_rows) <- names(col_colors) <- group_cols

  # Identify var block boundaries from spread structure:
  #   each block starts at a row where item is non-NA and non-empty,
  #   and ends with a totals row where both item and var are NA.
  item_col   <- spread[["item"]]
  label_rows <- which(!is.na(item_col) & item_col != "")
  totals_idx <- which(is.na(spread[["item"]]) & is.na(spread[["var"]]))

  current_excel_row <- 3L

  for (k in seq_along(label_rows)) {
    block_start    <- label_rows[k]
    totals_row_idx <- totals_idx[k]
    data_rows_idx  <- seq.int(block_start, totals_row_idx - 1L)
    n_data_rows    <- length(data_rows_idx)

    # Column totals stored in the totals row (absolute group-level sample sizes)
    totals_vec <- unlist(spread[totals_row_idx, group_cols, with = FALSE])

    for (grp_label in group_labels) {
      grp_prefix    <- paste0("[", grp_label, "] ")
      grp_col_names <- group_cols[startsWith(group_cols, grp_prefix)]
      if (length(grp_col_names) == 0L) next

      grp_totals <- as.numeric(totals_vec[grp_col_names])
      if (any(is.na(grp_totals)) || sum(grp_totals, na.rm = TRUE) == 0) next

      # Recover raw counts: n_ij = col_proportion_ij × col_total_j
      props_mat  <- as.matrix(spread[data_rows_idx, grp_col_names, with = FALSE])
      counts_mat <- sweep(props_mat, 2L, grp_totals, `*`)

      # Compute adjusted standardized residuals inline.
      # The original compute_adj_residuals_single accidentally included the
      # group-totals row in the count matrix, which doubles the column totals
      # and the grand total but halves the effective row proportions.  We
      # replicate that behaviour here (by appending colSums as an extra row)
      # so that the residual values — and therefore cell colours — are
      # numerically identical to the previous implementation.
      counts_ext  <- rbind(counts_mat, colSums(counts_mat))
      row_totals  <- rowSums(counts_ext)
      col_totals_ <- colSums(counts_ext)
      grand_total <- sum(counts_ext)

      if (grand_total <= 0 || !is.finite(grand_total)) next

      expected  <- outer(row_totals, col_totals_) / grand_total
      row_props <- row_totals / grand_total
      col_props <- col_totals_ / grand_total
      adj_resid_ext <- (counts_ext - expected) /
        sqrt(expected * outer(1 - row_props, 1 - col_props))
      # Drop the extra totals row — keep only the per-level residuals
      adj_resid <- adj_resid_ext[seq_len(nrow(counts_mat)), , drop = FALSE]

      excel_rows <- seq.int(current_excel_row, current_excel_row + n_data_rows - 1L)

      for (j in seq_along(grp_col_names)) {
        col_name <- grp_col_names[j]
        colors   <- residual_to_color_vec(adj_resid[, j])
        colored  <- which(!is.na(colors))
        if (length(colored) > 0L) {
          col_rows[[col_name]]   <- c(col_rows[[col_name]],   excel_rows[colored])
          col_colors[[col_name]] <- c(col_colors[[col_name]], colors[colored])
        }
      }
    }

    current_excel_row <- current_excel_row + n_data_rows + 1L
  }

  # One pass per group column: apply run-length-encoded fills
  for (col_name in group_cols) {
    rows   <- col_rows[[col_name]]
    colors <- col_colors[[col_name]]
    if (length(rows) == 0L) next

    excel_col <- which(names(spread) == col_name)
    if (length(excel_col) != 1L) next

    wb <- apply_column_fills(wb, int_to_col(excel_col), rows, colors)
  }

  wb
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
#' get_categerical_vars(iris)
#'
#' # Include character variables
#' get_categerical_vars(mtcars, include_char = TRUE)
#'
#' # Exclude specific variables
#' get_categerical_vars(iris, exclude = c("Species"))
#'
#' @export
get_categerical_vars <- function(data, exclude = NULL, include_char = FALSE) {
  # Get names of factor variables
  factor_vars <- names(data)[sapply(data, is.factor)]

  # Optionally include character variables
  if (include_char) {
    char_vars <- names(data)[sapply(data, is.character)]
    factor_vars <- c(factor_vars, char_vars)
  }

  # Exclude specified variables
  if (!is.null(exclude)) {
    factor_vars <- setdiff(factor_vars, exclude)
  }

  factor_vars
}
