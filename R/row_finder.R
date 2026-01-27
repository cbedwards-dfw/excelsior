#' Find row number based on an "anchor cell"
#'
#' Excel sheets are sometimes updated to add additional rows. This function find rows of interest based of off "anchor cells" -- cells with known contents (e.g., column headers) in a known column. Using these instead of hard-coding row numbers can make automation less
#' fragile to addition of rows.
#'
#' @param master_file Character. Filepath to Excel file of interest
#' @param sheet Character or numeric. Sheet name or index to search in
#' @param column Character or numeric. Column letter (e.g., "A") or number to search for pattern
#' @param pattern Character. Regular expression pattern to match in specified column
#' @param instance Numeric. Which match to use if multiple matches found? Defaults to 1.
#' @param offset Numeric. Number to add to returned row number. Useful when pattern identifies a title row and you want the row below (1) or above (-1). Defaults to 0
#'
#' @return Numeric. Row number where pattern was found, adjusted by offset
#'
#' @examples
#' \dontrun{
#' row_finder("data.xlsx", "Sheet1", "A", "Total", instance = 1, offset = 1)
#' row_finder("data.xlsx", 1, 3, "Summary", offset = -1)
#' }
#'
#' @export
## Finds rows based off an "anchor" cell in specific column with known pattern.
row_finder <- function(master_file, ## filepath to file of interest
                       sheet, ## sheet of interest
                       column, ## column number or letter to look for pattern
                       pattern, ##pattern to match in `column` to identify row of interest
                       instance = 1, ##which match to use? defaults to 1
                       offset = 0){ ## add to the returned value. Useful when the pattern is a title row and you want to start one row below (1) or above (-1) from that.

  validate_character(master_file, n = 1)
  validate_character(sheet, n = 1)
  if(length(column) != 1)(
    cli::cli_abort("`column` must be a single column identifier, either the column number (e.g., `3`), or the excel column identifier (e.g., `\"C\"`).")
  )
  validate_character(pattern, n = 1)
  validate_integer(instance, n = 1)
  validate_integer(offset, n = 1)

  if(is.character(column)){
    col_labels = tidyr::expand_grid(c("", LETTERS), LETTERS) |>
      apply(MARGIN =1, FUN = function(x){paste(x, collapse = "")})
    if(! column %in% col_labels){
      cli::cli_abort("If providing column as a character, must be excel column identifiers of form `\"A\"`, `\"B\"`, \"BX\", etc.")
    }
    column = which(col_labels == column)
  }
  raw <- readxl::read_excel(master_file,
                    sheet = sheet,
                    col_names = FALSE,
                    .name_repair = "unique_quiet")
  grep(pattern, raw[[column]])[instance] + offset
}
