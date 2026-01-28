#' Directly implant data from dataframe to openxlsx2 workbook
#'
#' Implants a dataframe into a specified cell range into an openxlsx2 workbook after
#' validation checks have been completed.
#'
#' @param wb An openxlsx2 workbook object
#' @param new_dat A dataframe containing the new data to be implanted
#' @param cell_range Character string specifying the range of cells to copy into (e.g., "A1:C3")
#' @param sheet Character string or numeric index specifying the sheet to copy into
#' @param numeric_flag Logical indicating whether the data should be converted to numeric (default: TRUE)
#' @param debug_mode Logical indicating whether to highlight updated cells in red (default: FALSE)
#'
#' @return Modified workbook object with data implanted
#'
#' @details
#' When \code{numeric_flag} is TRUE, all columns are converted to numeric.
#' When \code{debug_mode} is TRUE, updated cells are highlighted in blood red (#880808)
#' and an alert message is printed.
#' @export
implant_df <- function(wb,
                       new_dat,
                       cell_range,
                       sheet,
                       numeric_flag = TRUE,
                       debug_mode = FALSE){ ## if TRUE, highlight the cells that were updated
  validate_data_frame(new_dat)
  validate_character(cell_range, n = 1)
  validate_character(sheet, n = 1)
  validate_flag(numeric_flag)
  validate_flag(debug_mode)

  if(numeric_flag){
    new_dat <- new_dat |>
      dplyr::mutate(dplyr::across(dplyr::everything(),
                                  as_numeric_smart))
  }

  res <- wb |>
    openxlsx2::wb_add_data(sheet = sheet,
                           x = new_dat,
                           dims = cell_range,
                           col_names = FALSE,
                           row_names = FALSE,
                           apply_cell_style = FALSE,
                           na = "")
  ## If in debug mode, print out the update that was made, and highlight the
  ##   cells in blood red (#880808)
  if(debug_mode){
    cli::cli_alert("Handling updates to '{sheet}' {cell_range}")
    res <- res |>
      openxlsx2::wb_add_fill(sheet = sheet,
                             dims = cell_range,
                             color = openxlsx2::wb_color(hex = "#880808"))
  }
  return(res)

}

