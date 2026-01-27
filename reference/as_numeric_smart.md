# Convert character vector to numeric with intelligent handling

Converts a character vector to numeric while intelligently handling
common formatting that appears when reading mixed numeric and character
cells from excel, like commas and percentages.

## Usage

``` r
as_numeric_smart(vec)
```

## Arguments

- vec:

  A vector to convert to numeric

## Value

A numeric vector with the following transformations applied:

- Commas are removed (e.g., "1,000" becomes 1000)

- Percentages are removed and those numbers are converted to proportions
  (e.g., "10%" becomes 0.1)

- Values that cannot be converted return NA

## Examples

``` r
as_numeric_smart(c("1,000", "10%", "5.5", "text"))
#> [1] 1000.0    0.1    5.5     NA
# Returns: c(1000, 0.1, 5.5, NA)
```
