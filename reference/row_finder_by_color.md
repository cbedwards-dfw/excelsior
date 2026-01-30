# Find anchor cell based on fill color

Useful for handling sheets that use color-filled rows as separators
(e.g., STT table files).

## Usage

``` r
row_finder_by_color(
  filepath,
  sheet,
  column = 1,
  color = "FF000000",
  instance = 1,
  offset = 0
)
```

## Arguments

- filepath:

  Character. Filepath to Excel file of interest

- sheet:

  Character or numeric. Sheet name or index to search in

- column:

  Character or numeric. Column that contains anchor cell as number
  (e.g., 3) or corresponding excel column label ("C").

- color:

  Character vector. Color of anchor cell in hex, with or without
  preliminary "#". Defaults to black ("FF000000").

- instance:

  Numeric. Which match to use if multiple matches found? Defaults to 1.

- offset:

  Numeric. Number to add to returned row number. Useful when pattern
  identifies a title row and you want the row below (1) or above (-1).
  Defaults to 0

## Value

Row number (numeric of length 1)

## Examples

``` r
if (FALSE) { # \dontrun{
## find the row before the first column has a cell with black fill
row_finder_by_color("Appendix D.xlsx", sheet = "D-3", column = 1, offset = -1)
} # }
```
