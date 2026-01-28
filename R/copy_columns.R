#' Copy columns from one Excel file to another with change tracking
#'
#' Copies specified columns from an input Excel file to a workbook, highlighting
#' changes compared to a master file and adding metadata about the transfer. `row_start` and `row_end` are most easily identified using [row_finder()].
#'
#' @param wb An openxlsx2 workbook object to modify
#' @param to_file Character. Filepath for the "original" reference file
#' @param from_file Character. Filepath containing the workbook with data to copy FROM
#' @param sheet Character. Sheet name to copy from/to
#' @param columns Character vector. Excel column letters to copy (e.g., c("A", "B"))
#' @param row_start Integer. Initial row number to copy
#' @param row_end Integer. Final row number to copy
#' @param color_scheme Character vector of length 2. Hex colors for "Changed" and "Unchanged" cells. Use the same general color with different shades.
#' @param meta_cell Character. Excel cell reference to add transfer metadata (color, file, date).
#' @param verbose Logical. Whether to print progress messages. Defaults to TRUE.
#' @seealso [row_finder()]
#'
#' @return Modified workbook object with copied data, change highlighting, and metadata
#'
#' @export
#'
#' @examples
#' \dontrun{
#' wb <- load_workbook("primary.xlsx)
#' wb <- copy_columns(wb,
#'                   to_file = "primary.xlsx",
#'                   from_file = "update.xlsx",
#'                   sheet = "Sheet1",
#'                   columns = c("A", "B"),
#'                   row_start = 2,
#'                   row_end = 100,
#'                   color_scheme = c("#82DE4E", "#D5F4C3"),
#'                   meta_cell =  "D1")
#' }
copy_columns = function(wb,
                        to_file,
                        from_file,
                        sheet,
                        columns,
                        row_start,
                        row_end,
                        color_scheme,
                        meta_cell,
                        verbose = TRUE){
  from_file_short = glue::glue("{basename(dirname(from_file))}/{basename(from_file)}")
  if(verbose){
    cli::cli_h1("Transfering from {from_file_short}")
  }
  for(i_col in 1:length(columns)){
    cell_range = glue::glue("{columns[i_col]}{row_start}:{columns[i_col]}{row_end}")
    if(verbose){
      cli::cli_alert("Updating {cell_range} in sheet {sheet}")
    }

    dat_in = readxl::read_excel(from_file,
                        sheet = sheet,
                        range = cell_range,
                        col_names = FALSE,
                        .name_repair = "unique_quiet")
    dat_original = readxl::read_excel(to_file,
                              sheet = sheet,
                              range = cell_range,
                              col_names = FALSE,
                              .name_repair = "unique_quiet")
    dat_changed <- dat_in != dat_original
    dat_changed = dplyr::coalesce(dat_changed, FALSE)


    names(dat_in)[1] = "dat"
    dat_in$ind = seq.int(nrow(dat_in))
    dat_in$changed = as.vector(dat_changed)
    dat_in$cell = glue::glue("{columns[i_col]}{row_start:row_end}")
    suppressWarnings({
      dat_in <- dat_in |>
        dplyr::mutate(color = dplyr::if_else(.data$changed,
                               color_scheme[1],
                               color_scheme[2])) |>
        dplyr::mutate(numeric_check = !is.na(as_numeric_smart(.data$dat)))
    })

    dat_update <- dat_in |>
      dplyr::filter(.data$changed == TRUE) |>
      dplyr::mutate(
        # Extract column letter and row number
        col = stringr::str_extract(.data$cell, "^[A-Z]+"),
        row = as.integer(stringr::str_extract(.data$cell, "[0-9]+$")),
        # Create group ID for contiguous cells with same changed/check_numeric
        group_id = cumsum(.data$numeric_check != dplyr::lag(.data$numeric_check, default = !.data$numeric_check[1]) |
                            .data$row != dplyr::lag(.data$row, default = .data$row[1] - 1) + 1)
      ) |>
      dplyr::group_by(.data$group_id, .data$numeric_check)  |>
      dplyr::summarise(
        start_cell = dplyr::first(.data$cell),
        end_cell = dplyr::last(.data$cell),
        start_ind = dplyr::first(.data$ind),
        end_ind = dplyr::last(.data$ind),
        cell_range = dplyr::if_else(dplyr::first(.data$cell) == dplyr::last(.data$cell),
                             dplyr::first(.data$cell),
                             paste0(dplyr::first(.data$cell), ":", dplyr::last(.data$cell))),
        .groups = "drop"
      ) |>
      dplyr::select(-.data$start_cell, -.data$end_cell)

    if(nrow(dat_update)>0){
      for(i_row in 1:nrow(dat_update)){
        ## cutting from dataframe instead of reloading from excel for speed
        dat_cur = dat_in$dat[dat_update$start_ind[i_row]:dat_update$end_ind[i_row]]
        if(dat_update$numeric_check[i_row]){
          dat_cur = as_numeric_smart(dat_cur) |>
            round(digits = 10)
        }
        wb = wb |>
          openxlsx2::wb_add_data(sheet = sheet,
                      dat_cur,
                      dims = dat_update$cell_range[i_row],
                      col_names = FALSE,
                      na = "")
      }
    }

    ## handle colors efficiently
    dat_colors <- dat_in |>
      dplyr::filter(!is.na(.data$dat)) |>
      dplyr::mutate(
        # Extract column letter and row number
        col = stringr::str_extract(.data$cell, "^[A-Z]+"),
        row = as.integer(stringr::str_extract(.data$cell, "[0-9]+$")),
        # Create group ID for contiguous cells with same changed/check_numeric
        group_id = cumsum(.data$changed != dplyr::lag(.data$changed, default = !.data$changed[1]) |
                            row != dplyr::lag(.data$row, default = .data$row[1] - 1) + 1)
      )  |>
      dplyr::group_by(.data$group_id, .data$color) |>
      dplyr::summarise(
        start_cell = dplyr::first(.data$cell),
        end_cell = dplyr::last(.data$cell),
        cell_range = dplyr::if_else(dplyr::first(.data$cell) == dplyr::last(.data$cell),
                             dplyr::first(.data$cell),
                             paste0(dplyr::first(.data$cell), ":", dplyr::last(.data$cell))),
        .groups = "drop"
      )
    for(i_color in 1:nrow(dat_colors)){
      wb <- wb |>
        openxlsx2::wb_add_fill(sheet = sheet,
                    dims = dat_colors$cell_range[i_color],
                    color = openxlsx2::wb_color(hex = dat_colors$color[i_color]))
    }
  }

  ## Add metadata
  meta_dat = glue::glue("From `{from_file_short}`, copied in {date()}")
  wb <- wb |>
    openxlsx2::wb_add_data(sheet = sheet,
                meta_dat,
                dims = meta_cell,
                col_names = FALSE) |>
    openxlsx2::wb_add_fill(sheet = sheet,
                dims = meta_cell,
                color = openxlsx2::wb_color(hex = color_scheme[1]))

  return(wb)
}
