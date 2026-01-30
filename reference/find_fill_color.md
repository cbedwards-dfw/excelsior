# Find fill color of specified cell

Find fill color of specified cell

## Usage

``` r
find_fill_color(filepath, sheet, address)
```

## Arguments

- filepath:

  Excel file

- sheet:

  Excel sheet

- address:

  Cell address

## Value

Excel rgb color (character vector of length 1). To use in R, add "#" to
beginning.

## Examples

``` r
if (FALSE) { # \dontrun{
find_fill_color(primary_file, "D-3", "A103")
} # }
```
