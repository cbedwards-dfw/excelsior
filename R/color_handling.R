#' Find anchor cell based on fill color
#'
#' Useful for handling sheets that use color-filled rows as separators (e.g., STT table files).
#'
#' @inheritParams row_finder
#' @param color Character vector. Color of anchor cell in hex, with or without preliminary "#". Defaults to black ("FF000000").
#'
#' @return Row number (numeric of length 1)
#' @export
#'
#' @examples
#' \dontrun{
#' ## find the row before the first column has a cell with black fill
#' row_finder_by_color("Appendix D.xlsx", sheet = "D-3", column = 1, offset = -1)
#' }
row_finder_by_color <- function(filepath,
                                sheet,
                                column = 1,
                                color = "FF000000",
                                instance = 1,
                                offset = 0){ ## black

  validate_character(filepath, n = 1)
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

  ## validating color, then removing #
  color <- gsub("^[#]", "", color)
  color <- paste0("#", color)
  validate_hex_color(color, n = 1)
  color <- gsub("^[#]", "", color)

  validate_integer(instance, n = 1)
  validate_integer(offset, n = 1)

  if(is.character(column)){
    column = openxlsx2::col2int(column)
  }


  dat <- tidyxl::xlsx_cells(filepath, sheets = sheet)
  formats <- tidyxl::xlsx_formats(filepath)

  styles_want <- which(formats$local$fill$patternFill$fgColor$rgb == color)
  row_matching <- dat |>
    dplyr::filter(.data$local_format_id %in% .env$styles_want) |>
    dplyr::pull(.data$row)
  if(length(row_matching) == 0){
    cli::cli_abort("No cells in this column have the specified color ({.val {color}})! Consider using {.fun find_color} on anchor cell to identify color.")
  } else{
    row_matching = sort(unique(row_matching))
    return(row_matching[instance]+offset)
  }
}




#' Find fill color of specified cell
#'
#' @param filepath Excel file
#' @param sheet Excel sheet
#' @param address Cell address
#'
#' @return Excel rgb color (character vector of length 1). To use in R, add "#" to beginning.
#' @export
#'
#' @examples
#' \dontrun{
#' find_fill_color(primary_file, "D-3", "A103")
#' }
find_fill_color <- function(filepath, sheet, address){
  validate_character(filepath, 1)
  validate_character(sheet, 1)
  validate_cell_address(address, 1)

  dat <- tidyxl::xlsx_cells(filepath, sheets = sheet)
  formats <- tidyxl::xlsx_formats(filepath)

  cell_ind <- dat |>
    dplyr::filter(.data$address == .env$address) |>
    dplyr::pull(.data$local_format_id)

  color = formats$local$fill$patternFill$fgColor$rgb[cell_ind]
  return(color)
}

