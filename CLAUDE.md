# spreadview — developer orientation

## What the package does

Exports survey data as styled Excel spreadsheets. The main user entry point is
`compose_spreadsheet()`, which produces a workbook with one frequency table per
variable, optionally cross-tabulated by grouping variables and colored by
adjusted standardized residuals.

## Pipeline

```
compose_spreadsheet()          # orchestrates everything
  ├── check_input()            # validate data/vars/group
  ├── compose_table()          # per variable → frequency table (data.table)
  │     └── count_freqs()      # core frequency/proportion calculation
  ├── mark_percentage_cols()   # tag proportion columns with class "percentage"
  ├── compose_headers()        # two-row merged header in the workbook
  ├── compose_residuals()      # cell fill colors from adjusted residuals
  └── style_spreadsheet()      # column widths, number formats, borders
```

`compute_adj_residuals()` is also public and can be called standalone to get a
residual table without producing an Excel file.

## File map

### Public API — one exported function per file

| File | Exported function |
|---|---|
| `R/compose_spreadsheet.R` | `compose_spreadsheet()` |
| `R/compose_table.R` | `compose_table()` |
| `R/count_freqs.R` | `count_freqs()` |
| `R/compute_adj_residuals.R` | `compute_adj_residuals()` |
| `R/get_categorical_vars.R` | `get_categorical_vars()` |

### Internal helpers — no exports, never mixed with public functions

| File | Contents |
|---|---|
| `R/validate.R` | `check_input()`, `check_weights()` |
| `R/excel_utils.R` | workbook building and styling helpers |
| `R/residuals_internal.R` | `compute_adj_residuals_single()`, `residual_to_color_vec()` |
| `R/string_utils.R` | `extract_pattern()`, `extract_pattern_outside()` |

## Key conventions

- **data.table throughout.** All frequency calculations use `data.table`. Input
  data frames are converted internally. Use `data.table::set()` for in-place
  column mutation; avoid `<-` on columns inside `[.data.table]`.

- **Grouped column naming.** When a grouping variable is present, columns are
  named `[group_label] level`, e.g. `[island] Biscoe`. This pattern is used
  everywhere — in `count_freqs()` output, `compose_table()` output, and all
  Excel helpers. Grep for `"^\\[.+\\]"` to identify grouped columns.

- **Totals row.** Each variable block ends with a totals row where both `item`
  and `var` are `NA`. Grouped columns in this row hold absolute counts (not
  proportions), used by `compose_residuals()` to recover raw cell counts.

- **No re-exports.** Internal helpers are not exported. Access them in
  development via `spreadview:::fn_name()` or by loading with `devtools::load_all()`.

## Adding a new variable type

1. Extend `check_input()` in `R/validate.R` to accept the new type.
2. Add a branch in `count_freqs()` in `R/count_freqs.R` for the new
   frequency calculation logic.
3. Update `compose_table()` in `R/compose_table.R` if the table structure
   needs to change.
4. Add tests in `tests/testthat/test-count_freqs.R`.

## Adding a new styling rule

All styling lives in `R/excel_utils.R`. Add the new rule inside
`style_spreadsheet()`. Use `openxlsx2::wb_dims()` to address cell ranges and
the existing border/format patterns as a template.

## Running checks

```r
devtools::load_all()   # load package for interactive testing
devtools::test()       # run all tests
devtools::check()      # full R CMD check
devtools::document()   # regenerate NAMESPACE and man/ after roxygen changes
```
