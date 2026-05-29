expand_cell_range <- function(range) {
  # Split on ":" to get the two corner addresses
  corners <- strsplit(range, ":")[[1]]
  if (length(corners) != 2) stop("Input must be a range in 'A1:B2' format.")

  # Helper: split a cell address into column letters and row number
  parse_cell <- function(cell) {
    col_str <- gsub("[0-9]", "", cell)
    row_num <- as.integer(gsub("[^0-9]", "", cell))
    list(col = col_str, row = row_num)
  }

  # Helper: convert column letters to a 1-based integer (e.g. "A"->1, "Z"->26, "AA"->27)
  col_to_int <- function(col) {
    chars <- strsplit(toupper(col), "")[[1]]
    Reduce(function(acc, ch) acc * 26L + utf8ToInt(ch) - 64L, chars, 0L)
  }

  # Helper: convert a 1-based integer back to column letters
  int_to_col <- function(n) {
    result <- ""
    while (n > 0) {
      remainder <- (n - 1) %% 26
      result <- paste0(intToUtf8(remainder + 65L), result)
      n <- (n - 1) %/% 26
    }
    result
  }

  top_left  <- parse_cell(corners[1])
  bot_right <- parse_cell(corners[2])

  col_start <- col_to_int(top_left$col)
  col_end   <- col_to_int(bot_right$col)
  row_start <- top_left$row
  row_end   <- bot_right$row

  # Build every combination of column x row
  cols <- vapply(seq(col_start, col_end), int_to_col, character(1))
  rows <- seq(row_start, row_end)

  as.vector(outer(cols, rows, paste0))
}
