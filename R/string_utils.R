# Internal string utilities for parsing variable labels and codes.
# These are not part of the public API.

# Extract substrings matching a pattern from a character vector.
# Returns the original element for non-matching entries.
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

# Remove a pattern (and trailing whitespace) from a character vector.
# Returns an empty string for non-matching entries.
extract_pattern_outside <- function(x, pattern) {
  if (!is.character(x)) {
    stop("x has to be character vector")
  }

  result <- sub(pattern = paste0(pattern, "\\s*"), replacement = "", x)
  ifelse(result == x, "", result)
}
