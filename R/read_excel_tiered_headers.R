## Helper function
collapse_helper = function(x, sep_char = "_"){
  x = stats::na.omit(x)
  if(length(x) > 1){
    is_duped = (x[-1] == x[-length(x)])
    ## first occurence can't be duped
    is_duped = c(FALSE, is_duped)
    x = x[!is_duped]
  }
  paste0(x, collapse = sep_char)
}

#' Read an Excel file with multi-row headers
#'
#' Reads an Excel worksheet where column headers are spread across multiple
#' rows, including merged cells. Header rows are combined into a single header
#' string per column, with consecutive duplicate values collapsed (e.g. a
#' superheader "TREATY" spanning columns that are individually labelled
#' "TROLL", "NET", "SPORT" becomes "TREATY_TROLL", "TREATY_NET",
#' "TREATY_SPORT").
#'
#' Merged cells are handled via [openxlsx2::read_xlsx()] with
#' `fill_merged_cells = TRUE`, which propagates the anchor cell's value
#' across the full span of the merge before any header combination is
#' attempted.
#'
#' @param path Character. Path to the \code{.xlsx} file.
#' @param sheet Integer or character atomic Sheet index to read. Default \code{1} to automatically handle cases with a single sheet.
#' @param header_rows Integer vector. Row numbers (1-indexed, as they appear
#'   in the spreadsheet) that together form the column headers.
#' @param pseudo_merged_rows Integer vector. Row numbers (1-indexed, as they appear
#'   in the spreadsheet) for header rows that include "false merged cells" -- cells that look like they're merged because the first cell of a group has weird formatting that shifts the text far to the right. Warning: this can create funky names cases in which there are gaps columns in the merged row. This is NOT necessary if the excel cells are *actually* merged. Defaults to NULL.
#' @param first_data_row Integer or \code{NULL}. Row number of the first data
#'   row. If \code{NULL} (default), set to \code{max(header_rows) + 1}.
#' @param final_data_row Integer or \code{NULL}. Row number of the last data
#'   row. If \code{NULL} (default), all rows from \code{first_data_row} to
#'   the end of the sheet are returned.
#' @param sep Character. Separator used when joining header parts from
#'   different rows. Default \code{"_"}.
#' @param clean_names Logical. Resolves issues in which column names are not unique.  If \code{TRUE} (default), column names are
#'   passed through [base::make.names()] with \code{unique = TRUE} to ensure
#'   they are valid and non-duplicate R names.
#'
#' @return A data frame with one column per spreadsheet column and one row per
#'   data row. Column names are the combined headers derived from
#'   \code{header_rows}.
#'
#' @export
#'
#' @examples
#' \dontrun{
#' # Basic usage: rows 2-4 are headers, data starts at row 5
#' df <- read_excel_tiered_headers("data.xlsx")
#'
#' # Custom header rows and separator
#' df <- read_excel_tiered_headers(
#'   path        = "data.xlsx",
#'   header_rows = c(1, 2),
#'   sep         = "."
#' )
#'
#' # Read only a specific range of data rows
#' df <- read_excel_tiered_headers(
#'   path           = "data.xlsx",
#'   hearder_rows = c(2, 3, 4)
#'   first_data_row = 7,
#'   final_data_row = 100
#' )
#' }
read_excel_tiered_headers <- function(path,
                                      sheet = 1,
                                      header_rows,
                                      pseudo_merged_rows = NULL,
                                      first_data_row = NULL,
                                      final_data_row = NULL,
                                      sep = "_",
                                      clean_names = TRUE){

  if (is.null(first_data_row)){
    first_data_row <- max(header_rows) + 1
  }

  ## read in, respecting merged cells
  raw <- openxlsx2::read_xlsx(path,
                              sheet = sheet,
                              start_row = 1,
                              fill_merged_cells = TRUE,
                              col_names = FALSE)

  if(!is.null(pseudo_merged_rows)){
    for(cur_row in pseudo_merged_rows){
      raw[cur_row, ] = zoo::na.locf(as.vector(raw[cur_row, ]), na.rm = FALSE)
    }
  }

  if(is.null(final_data_row)){
    final_data_row = nrow(raw)
  }

  ## initial cleaning
  headers_clean <- apply(raw[header_rows, ], 2, collapse_helper, sep_char = sep)

  if (clean_names){
    headers_clean <- make.names(headers_clean, unique = TRUE)
  }

  data <- openxlsx2::read_xlsx(path,
                               sheet = sheet,
                               rows = first_data_row:final_data_row,
                               fill_merged_cells = TRUE,
                               col_names = FALSE)
  names(data) <- headers_clean[1:ncol(data)]

  return(data)
}

