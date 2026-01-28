#' Find row number based on an "anchor cell"
#'
#' Excel sheets are sometimes updated to add additional rows. This function find rows of interest based of off "anchor cells" -- cells with known contents (e.g., column headers) in a known column. Using these instead of hard-coding row numbers can make automation less
#' fragile to addition of rows.
#'
#' The `pattern` argument accepts regular expressions, which are more flexible than just exactly matching a string. A few tips:
#'
#'  - R treats `\` specially -- if present in the pattern, surround with square brackets to treat them as actual characters.
#'  - In regular expressions, the `^` symbol is used to denote the start of a string, and the `$` symbol is used to denote the end of a string.
#'
#'  As an example of this in practice, if we had an anchor cell that was the footnote "a/  Derived from vessel registrations and fish landing tickets.", we could use as a pattern any section of the text (e.g., `pattern = "Derived from vessel registrations and fish"`). But if we wanted to ensure that the first occurrence of "a/..." was used as our anchor (in case the footnote text changes), we could use the pattern `"^a[/]"`, which in words means "the first letter is an a, and the second letter is literally a `/`".
#'
#' @param master_file Character. Filepath to Excel file of interest
#' @param sheet Character or numeric. Sheet name or index to search in
#' @param column Character or numeric. Column that contains anchor cell as number (e.g., 3) or corresponding excel column label ("C").
#' @param pattern Character. Pattern to identify anchor cell. Accepts regular expressions -- see details.
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
row_finder <- function(master_file,
                       sheet,
                       column,
                       pattern,
                       instance = 1,
                       offset = 0){

  validate_character(master_file, n = 1)
  validate_character(sheet, n = 1)
  ## validating column
  if(!is.numeric(column) & !is.character(column)){
    cli::cli_abort("`column` must identify a single column using either column number (e.g., `3`), or the excel column identifier (e.g., `\"C\"`).")
  }
  if(is.numeric(column)){
    validate_integer(column)
  } else {
    validate_excel_column(column)
  }

  validate_character(pattern, n = 1)
  validate_integer(instance, n = 1)
  validate_integer(offset, n = 1)

  if(is.character(column)){
    column = openxlsx2::col2int(column)
  }

  raw <- readxl::read_excel(master_file,
                    sheet = sheet,
                    col_names = FALSE,
                    .name_repair = "unique_quiet")
  grep(pattern, raw[[column]])[instance] + offset
}
