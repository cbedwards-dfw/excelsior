# Copy columns from one Excel file to another with change tracking

Copies specified columns from an input Excel file to a workbook,
highlighting changes compared to a master file and adding metadata about
the transfer. `row_start` and `row_end` are most easily identified using
[`row_finder()`](https://cbedwards-dfw.github.io/excelsior/reference/row_finder.md).

## Usage

``` r
copy_columns(
  wb,
  master_file,
  in_file,
  sheet,
  columns,
  row_start,
  row_end,
  color_scheme,
  meta_cell,
  verbose = TRUE
)
```

## Arguments

- wb:

  An openxlsx2 workbook object to modify

- master_file:

  Character. Filepath for the "original" reference file

- in_file:

  Character. Filepath containing the workbook with data to copy FROM

- sheet:

  Character. Sheet name to copy from/to

- columns:

  Character vector. Excel column letters to copy (e.g., c("A", "B"))

- row_start:

  Integer. Initial row number to copy

- row_end:

  Integer. Final row number to copy

- color_scheme:

  Character vector of length 2. Hex colors for "Changed" and "Unchanged"
  cells. Use the same general color with different shades.

- meta_cell:

  Character. Excel cell reference to add transfer metadata (color, file,
  date).

- verbose:

  Logical. Whether to print progress messages. Defaults to TRUE.

## Value

Modified workbook object with copied data, change highlighting, and
metadata

## See also

[`row_finder()`](https://cbedwards-dfw.github.io/excelsior/reference/row_finder.md)

## Examples
