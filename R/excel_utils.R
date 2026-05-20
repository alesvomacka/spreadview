# Internal helpers for building and styling the openxlsx2 workbook.
# None of these are part of the public API.


# Convert a 1-based column index to an Excel column letter (supports up to ZZ).
int_to_col <- function(n) {
  if (n <= 26L) return(LETTERS[n])
  paste0(LETTERS[(n - 1L) %/% 26L], LETTERS[(n - 1L) %% 26L + 1L])
}


# Mark columns from "total" onwards with class "percentage" so that
# style_spreadsheet() and compose_spreadsheet() can identify them later.
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

  header_df <- data.frame(
    rbind(header_row1, header_row2),
    stringsAsFactors = FALSE
  )
  names(header_df) <- col_names

  wb <- wb |>
    openxlsx2::wb_add_data(x = header_df, col_names = FALSE, start_row = 1)

  if (any(is_grouped)) {
    run_start <- NULL
    run_value <- NULL

    for (i in seq_along(header_row1)) {
      current_val <- header_row1[i]
      is_group_col <- is_grouped[i]

      if (is_group_col) {
        if (is.null(run_start) || current_val != run_value) {
          if (!is.null(run_start) && (i - 1) > run_start) {
            wb <- wb |>
              openxlsx2::wb_merge_cells(
                dims = openxlsx2::wb_dims(rows = 1, cols = run_start:(i - 1))
              )
          }
          run_start <- i
          run_value <- current_val
        }
      } else {
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


#' @keywords internal
compose_residuals <- function(wb, spread) {
  group_cols <- grep("^\\[.+\\]", names(spread), value = TRUE)
  if (length(group_cols) == 0L) return(wb)

  group_labels <- unique(sub("^\\[(.+?)\\] .+$", "\\1", group_cols))

  col_rows   <- vector("list", length(group_cols))
  col_colors <- vector("list", length(group_cols))
  names(col_rows) <- names(col_colors) <- group_cols

  item_col   <- spread[["item"]]
  label_rows <- which(!is.na(item_col) & item_col != "")
  totals_idx <- which(is.na(spread[["item"]]) & is.na(spread[["var"]]))

  current_excel_row <- 3L

  for (k in seq_along(label_rows)) {
    block_start    <- label_rows[k]
    totals_row_idx <- totals_idx[k]
    data_rows_idx  <- seq.int(block_start, totals_row_idx - 1L)
    n_data_rows    <- length(data_rows_idx)

    totals_vec <- unlist(spread[totals_row_idx, group_cols, with = FALSE])

    for (grp_label in group_labels) {
      grp_prefix    <- paste0("[", grp_label, "] ")
      grp_col_names <- group_cols[startsWith(group_cols, grp_prefix)]
      if (length(grp_col_names) == 0L) next

      grp_totals <- as.numeric(totals_vec[grp_col_names])
      if (any(is.na(grp_totals)) || sum(grp_totals, na.rm = TRUE) == 0) next

      props_mat  <- as.matrix(spread[data_rows_idx, grp_col_names, with = FALSE])
      counts_mat <- sweep(props_mat, 2L, grp_totals, `*`)

      row_totals  <- rowSums(counts_mat)
      col_totals_ <- colSums(counts_mat)
      grand_total <- sum(counts_mat)

      if (grand_total <= 0 || !is.finite(grand_total)) next

      expected  <- outer(row_totals, col_totals_) / grand_total
      row_props <- row_totals / grand_total
      col_props <- col_totals_ / grand_total
      adj_resid <- (counts_mat - expected) /
        sqrt(expected * outer(1 - row_props, 1 - col_props))

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


#' @keywords internal
style_spreadsheet <- function(wb, spread, group = NULL, percent = TRUE) {
  n_rows <- nrow(spread)
  n_cols <- ncol(spread)

  wb$set_col_widths(cols = 1, widths = 10)
  wb$set_col_widths(cols = 2, widths = 26)
  wb$set_col_widths(cols = 3, widths = 26)
  wb$set_col_widths(cols = 4:n_cols, widths = 10)

  n_col <- which(names(spread) == "n")
  dims_n <- openxlsx2::wb_dims(rows = seq(3, n_rows + 2), cols = n_col)
  wb$add_numfmt(dims = dims_n, numfmt = "0")

  N_col <- which(names(spread) == "N")
  if (length(N_col) > 0) {
    dims_N <- openxlsx2::wb_dims(rows = seq(3, n_rows + 2), cols = N_col)
    wb$add_numfmt(dims = dims_N, numfmt = "0")
  }

  total_col <- which(names(spread) == "total")
  if (length(total_col) > 0) {
    pct_cols <- seq(total_col, n_cols)
    dims_pct <- openxlsx2::wb_dims(
      rows = seq(3, n_rows + 2),
      cols = pct_cols
    )
    wb$add_numfmt(dims = dims_pct, numfmt = if (percent) "0%" else "0")
  }

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

  dims_merged_header <- openxlsx2::wb_dims(rows = 1:2, cols = 1:n_cols)
  wb$add_cell_style(
    dims = dims_merged_header,
    wrap_text = TRUE,
    horizontal = "center",
    vertical = "top"
  )

  dims_values <- openxlsx2::wb_dims(rows = 3:(n_rows + 2), cols = 4:n_cols)
  wb$add_cell_style(dims = dims_values, horizontal = "center")

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
