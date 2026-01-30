
validate_data_frame <- function(x, ..., arg = rlang::caller_arg(x), call = rlang::caller_env()) {
  # checks for data frame, stolen from the tidyr package
  if (!is.data.frame(x)) {
    cli::cli_abort("{.arg {arg}} must be a data frame, not {.obj_type_friendly {x}}.", ..., call = call)
  }
}


validate_numeric <- function(x, n = NULL, min = NULL, max = NULL, ..., arg = rlang::caller_arg(x), call = rlang::caller_env()) {
  if (!is.numeric(x)) {
    cli::cli_abort("{.arg {arg}} must be a numeric, not {class(x)}.", ..., call = call)
  }
  if (!is.null(n)) {
    if (length(x) != n) {
      cli::cli_abort("{.arg {arg}} must be a numeric of length {.val {n}}, not {.val {length(x)}}.", ..., call = call)
    }
  }
  if (!is.null(min)) {
    if (any(x < min)) {
      if (!is.null(n) && n > 1) {
        cli::cli_abort("All values of {.arg {arg}} must be no less than {min}.", ..., call = call)
      } else {
        cli::cli_abort("{.arg {arg}} must be no less than {.val {min}}.", ..., call = call)
      }
    }
  }
  if (!is.null(max)) {
    if (any(x > max)) {
      if (!is.null(n) && n > 1) {
        cli::cli_abort("All values of {.arg {arg}} must be no greater than than {max}.", ..., call = call)
      } else {
        cli::cli_abort("{.arg {arg}} must be no greater than than {.val {max}}.", ..., call = call)
      }
    }
  }
}

validate_character <- function(x, n = NULL, ..., arg = rlang::caller_arg(x), call = rlang::caller_env()) {
  if (!is.character(x)) {
    cli::cli_abort("{.arg {arg}} must be a character, not {.obj_type_friendly {x}}.", ..., call = call)
  }
  if(!is.null(n)){
    if(length(x) != n){
      cli::cli_abort("{.arg {arg}} must be a character of length {.val {n}}, not {.val {length(x)}}.", ..., call = call)
    }
  }
}

validate_flag <- function(x, ..., arg = rlang::caller_arg(x), call = rlang::caller_env()){
  if (!is.logical(x) | length(x) != 1) {
    cli::cli_abort("{.arg {arg}} must be a logical of length {.val 1}.", ..., call = call)
  }
}

validate_integer <- function(x, n = NULL, ..., arg = rlang::caller_arg(x), call = rlang::caller_env()) {
  validate_numeric(x, n, arg = arg, call = call, ...)
  if (any(x %% 1 != 0)) {
    if (!is.null(n) && n > 1) {
      cli::cli_abort("{.arg {arg}} must contain only whole numbers.", ..., call = call)
    } else {
      cli::cli_abort("{.arg {arg}} must be a whole number.", ..., call = call)
    }
  }
}

validate_wb <- function(x, ..., arg = rlang::caller_arg(x), call = rlang::caller_env()){
  if (! ("wbWorkbook" %in% class(x) )) {
    cli::cli_abort("{.arg {x}} must be an openxlsx2 workbook(a `wbWorkbook` object), not {.obj_type_friendly {x}}", ..., call = call)
  }
}

validate_excel_column <- function(x, n = NULL, ..., arg = rlang::caller_arg(x), call = rlang::caller_env()){
  validate_character(x, n = n, arg = arg, call = call)
  possible_cols = tidyr::expand_grid(c("", LETTERS), c("",LETTERS), LETTERS) |>
    apply(MARGIN =1, FUN = function(x){paste(x, collapse = "")})

  if(! all (x %in% possible_cols)) {
    cli::cli_abort("{.arg {x}} is not a legal excel column (allowed: {.val A} through {.val AFD}).", ..., call = call)
  }
}

validate_hex_color <- function(x, n = NULL, ..., arg = rlang::caller_arg(x), call = rlang::caller_env()) {
  # Check if it's a character vector
  validate_character(x, n = n, arg = arg, call = call)

  # Check if each element is a valid hex color
  # Pattern allows for #RGB, #RRGGBB, #RRGGBBAA formats
  hex_pattern <- "^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$"

  valid <- grepl(hex_pattern, x)

  if (!all(valid)) {
    invalid_indices <- which(!valid)
    invalid_values <- x[invalid_indices]
    cli::cli_abort("{.arg {arg}} must be {n} hex color code(s) (e.g., \"FF0000\"). Invalid values {.val {invalid_values}} at positions {.val {invalid_indices}}.", ..., call = call)
  }
}

validate_cell_address <- function(x, n = NULL, ..., arg = rlang::caller_arg(x), call = rlang::caller_env()) {
  validate_character(x, n = n, arg = arg, call = call)

  pattern <- "^[A-Z]+[0-9]+$"
  if (!all(grepl(pattern, x))) {
    cli::cli_abort("Elements of {.arg {arg}} must be valid Excel cell addresses (e.g., 'A1', 'B10', 'AA100').",
                   ..., call = call)
  }
}

validate_cell_range <- function(x, n = NULL, single_cell_allowed = TRUE, ..., arg = rlang::caller_arg(x), call = rlang::caller_env()) {
  validate_character(x, n = n, arg = arg, call = call)

  pattern <- "^[A-Z]+[0-9]+:[A-Z]+[0-9]+$"
  if(single_cell_allowed){
    pattern = paste0(pattern, "|^[A-Z]+[0-9]+$")
  }
  if (!all(grepl(pattern, x))) {
    if(single_cell_allowed){
      cli::cli_abort("Elements of {.arg {arg}} must be valid Excel cell ranges or cell address (e.g., 'D5', 'A1:B10', 'C5:AA100').",
                     ..., call = call)
    } else {
      cli::cli_abort("Elements of {.arg {arg}} must be valid Excel cell ranges (e.g., 'A1:B10', 'C5:AA100'). Single cell addresses (e.g., 'D5') are not allowed.",
                     ..., call = call)
    }
  }
}

validate_excel_sheet <- function(sheet, filepath, n = NULL, ..., arg = rlang::caller_arg(sheet), call = rlang::caller_env()) {
  validate_character(x = sheet, n = n)
  all_sheets <- readxl::excel_sheets(filepath)
  missing_sheets = setdiff(sheet, all_sheets)
  if(length(missing_sheets)>0){
    cli::cli_abort("{.arg {arg}} must contains sheets present in the excel file! {.val {missing_sheets}} not present in file. Available sheets: {.val {all_sheets}}.",
                   ..., call = call)
  }
}


# validate_cell_range("")

