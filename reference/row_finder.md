# Find row number based on an "anchor cell"

Excel sheets are sometimes updated to add additional rows. This function
find rows of interest based of off "anchor cells" – cells with known
contents (e.g., column headers) in a known column. Using these instead
of hard-coding row numbers can make automation less fragile to addition
of rows.

## Usage

``` r
row_finder(master_file, sheet, column, pattern, instance = 1, offset = 0)
```

## Arguments

- master_file:

  Character. Filepath to Excel file of interest

- sheet:

  Character or numeric. Sheet name or index to search in

- column:

  Character or numeric. Column letter (e.g., "A") or number to search
  for pattern

- pattern:

  Character. Regular expression pattern to match in specified column

- instance:

  Numeric. Which match to use if multiple matches found? Defaults to 1.

- offset:

  Numeric. Number to add to returned row number. Useful when pattern
  identifies a title row and you want the row below (1) or above (-1).
  Defaults to 0

## Value

Numeric. Row number where pattern was found, adjusted by offset

## Examples

``` r
if (FALSE) { # \dontrun{
row_finder("data.xlsx", "Sheet1", "A", "Total", instance = 1, offset = 1)
row_finder("data.xlsx", 1, 3, "Summary", offset = -1)
} # }
```
