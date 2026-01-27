#' Convert character vector to numeric with intelligent handling
#'
#' Converts a character vector to numeric while intelligently handling common
#' formatting that appears when reading mixed numeric and character cells from excel, like commas and percentages.
#'
#' @param vec A vector to convert to numeric
#'
#' @return A numeric vector with the following transformations applied:
#'   \itemize{
#'     \item Commas are removed (e.g., "1,000" becomes 1000)
#'     \item Percentages are removed and those numbers are converted to proportions (e.g., "10%" becomes 0.1)
#'     \item Values that cannot be converted return NA
#'   }
#'
#' @examples
#' as_numeric_smart(c("1,000", "10%", "5.5", "text"))
#' # Returns: c(1000, 0.1, 5.5, NA)
#'
#' @export
as_numeric_smart <- function(vec){
  result <- sapply(vec, function(x) {
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
    # if(!is.na(num_val)) {
    #   return(num_val)
    # }

    # Otherwise, keep as character
    return(num_val)
  }, USE.NAMES = FALSE)
  return(result)
}
