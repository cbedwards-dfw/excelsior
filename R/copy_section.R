#' Copy data section between locations in Excel workbook
#'
#' Designed with updating in mind. Copies data from one location to another within an Excel workbook, with optional validation checks for matching reference rows/columns.
#' This is a safe version of saying "Copy sheet C3:D4 of sheet 1 to E5:F6 of sheet 2".
#' To support speed, this function takes as arguments the openxlsx2 workbook object  to be updated, and a dataframe to update FROM.
#'
#' The key feature of this function is optional checking, in which the first column or row of the "from_db"
#' is checked against a corresponding column or row of the "to_address" to make sure the entries are still appropriate.
#'
#' As an example, when updating the TAMM template file with a new year's data, we have a rectangular section of locations and their values present in the TAMM and in the input template; we want to update the input template with the new TAMM values. We can code this with cell addresses, identifying that we want to copy from cells "D126:F130" of the TAMM to "D5:F9" of the current sheet of the input template. However, we could encounter problems if users have re-arranged either excel sheet. For this reason, we want to confirm that the location column (which is 1 column to the left of the leftmost cells we're copying) matches. So we provide "check_col_offset = 1" as an argument. This causes the function to check that "C123:C130" of the TAMM matches "C5:C9" of the current sheet of the input template.
#'
#' @param wb An openxlsx2 workbook object
#' @param from_df A dataframe of the complete sheet that contains the data to copy *from*. Typically generated from `readxl::read_excel` with `col_names = FALSE` and `.name_repair = "unique_quiet`.
#' @param from_address Character string specifying the source cell range (e.g., "A1:C3").
#' @param to_address Character string specifying the destination cell range (e.g., "D1:F3")
#' @param sheet Character string or numeric index specifying the target sheet
#' @param check_row_offset Optional, integer offset identifying the row that should be used to validate that the from/to sections match. Positive values correspond to rows *above* the first row of the addresses specified.
#' @param check_col_offset Optional, integer offset identifying the column that should be used to validate that the from/to sections match. Positive values correspond to columns *before* the first column of the addresses specified.
#' @param numeric_flag Logical indicating whether data should be converted to numeric (default: TRUE)
#' @param debug_mode Logical indicating whether to enable debug mode highlighting (default: FALSE)
#'
#' @return Modified workbook object with data copied
#'
#' @export
copy_section <- function(wb,
                         from_df,
                         from_address,
                         to_address,
                         sheet,
                         check_row_offset = NULL,
                         check_col_offset = NULL,
                         numeric_flag = TRUE,
                         debug_mode = FALSE){

  validate_character(from_address, n = 1)
  validate_character(to_address, n = 1)
  validate_character(sheet, n = 1)
  if(!is.null(check_row_offset)){validate_integer(check_row_offset, n = 1)}
  if(!is.null(check_col_offset)){validate_integer(check_col_offset, n = 1)}
  validate_flag(numeric_flag)
  validate_flag(debug_mode)

  ## turn from_address to dataframe of rows and columns
  df_dims = cell_range_translate(from_address, expand = FALSE)
  if(nrow(df_dims) == 1){
    df_dims = rbind(df_dims, df_dims)
  }

  ## check that from_address and to_address dimensions match
  df_dims_full = cell_range_translate(from_address)
  df_dims_to_full = cell_range_translate(to_address)
  if(length(unique(.data$df_dims_full$row)) != length(unique(.data$df_dims_to_full$row)) ||
     length(unique(.data$df_dims_full$col)) != length(unique(.data$df_dims_to_full$col))
  ){
    cli::cli_abort("Size of `from_address` and `to_address` must match!")
  }

  ## if offset row is present, check for matching references
  if(!is.null(check_row_offset)){
    ref_from <- from_df[.data$df_dims$row[1] - check_row_offset,
                        .data$df_dims$col[1]:.data$df_dims$col[2]] |>
      ## convert NAs to ""s for better comparing
      dplyr::mutate(dplyr::across(dplyr::everything(), \(x) dplyr::coalesce(x, "")))|>
      ## make sure leading/trailing strings are handled consistently
      dplyr::mutate(dplyr::across(dplyr::everything(), stringr::str_trim))
    df_dims_to = cell_range_translate(to_address, expand = FALSE)
    ## handling single-cell case
    if(nrow(df_dims_to) == 1){
      df_dims_to = rbind(df_dims_to, df_dims_to)
    }
    ref_to <- openxlsx2::wb_to_df(wb,
                                  sheet = sheet,
                                  col_names = FALSE)[.data$df_dims_to$row[1] - check_row_offset,
                                                     .data$df_dims_to$col[1]:.data$df_dims_to$col[2]]  |>
      tibble::tibble() |>
      ## convert NAs to ""s for better comparing
      dplyr::mutate(dplyr::across(dplyr::everything(), \(x) dplyr::coalesce(x, ""))) |>
      ## make sure leading/trailing strings are handled consistently
      dplyr::mutate(dplyr::across(dplyr::everything(), stringr::str_trim))
    if(!all(ref_from == ref_to)){
      cli::cli_abort("Reference mismatch for sheet \"{.strong {sheet}}\" when copying {.emph {from_address}} to {.emph {to_address}}! Look at {ref_from[ref_from != ref_to]}")
    }
  }


  ## if offset col is present, check for matching references
  if(!is.null(check_col_offset)){
    ref_from <- tibble::tibble(from_df[.data$df_dims$row[1]:.data$df_dims$row[2],
                                       .data$df_dims$col[1] - check_col_offset]) |>
      ## convert NAs to ""s for better comparing
      dplyr::mutate(dplyr::across(dplyr::everything(), \(x) dplyr::coalesce(x, "")))|>
      ## make sure leading/trailing strings are handled consistently
      dplyr::mutate(dplyr::across(dplyr::everything(), stringr::str_trim))
    df_dims_to = cell_range_translate(to_address, expand = FALSE)
    ## handling single-cell case
    if(nrow(df_dims_to) == 1){
      df_dims_to = rbind(df_dims_to, df_dims_to)
    }
    ref_to <- openxlsx2::wb_to_df(wb,
                                  sheet = sheet,
                                  col_names = FALSE)[.data$df_dims_to$row[1]:.data$df_dims_to$row[2],
                                                     .data$df_dims_to$col[1] - check_col_offset]  |>
      tibble::tibble() |>
      ## convert NAs to ""s for better comparing
      dplyr::mutate(dplyr::across(dplyr::everything(), \(x) dplyr::coalesce(x, ""))) |>
      ## make sure leading/trailing strings are handled consistently
      dplyr::mutate(dplyr::across(dplyr::everything(), stringr::str_trim))
    if(!all(ref_from == ref_to)){
      cli::cli_abort("Reference mismatch for sheet \"{.strong {sheet}}\" when copying {.emph {from_address}} to {.emph {to_address}}! Look at {ref_from[ref_from != ref_to]}")
    }
  }


  ## do the copying
  new_dat = from_df[(.data$df_dims$row[1]:.data$df_dims$row[2]),
                    .data$df_dims$col[1]:.data$df_dims$col[2]]

  implant_df(wb,
             new_dat = new_dat,
             cell_range = to_address,
             sheet = sheet,
             numeric_flag = numeric_flag,
             debug_mode = debug_mode)
}
