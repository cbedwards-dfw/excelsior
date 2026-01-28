#' Translate copied excel row/column into code to make an R vector
#'
#' Function interacts with the computer clipboard.
#' To use, start by copying a row or column in excel that you want coded as a
#' vector in R. Run clip_to_vec() to change the system clipboard to
#' the string of R code that defines an equivalent vector, then paste into
#' R script.
#'
#' @return Nothing
#' @export
#'
clip_to_vec <- function(){
  # Read from clipboard
  temp <- utils::read.delim("clipboard", sep = "\t", header = FALSE, stringsAsFactors = FALSE)

  # Convert to vector (handles both single row and single column)
  if(ncol(temp) == 1) {
    # Column input
    temp <- temp[[1]]
  } else {
    # Row input - convert to vector
    temp <- unlist(temp, use.names = FALSE)
  }

  # Process each element
  result <- sapply(temp, function(x) {
    # Skip if empty or NA
    if(is.na(x) || x == "") return(x)

    # Remove commas (for numbers like 1,000)
    x_clean <- gsub(",", "", x)

    # Check if it's a percentage
    if(grepl("%$", x_clean)) {
      return(as.numeric(sub("%$", "", x_clean)) / 100)
    }

    # Check if it's numeric (after removing commas)
    num_val <- suppressWarnings(as.numeric(x_clean))
    if(!is.na(num_val)) {
      return(num_val)
    }

    # Otherwise, keep as character
    return(x)
  }, USE.NAMES = FALSE)

  # Write vector definition to clipboard
  dput(result, "clipboard")
}
