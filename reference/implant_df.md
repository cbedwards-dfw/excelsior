# Directly implant data from dataframe to openxlsx2 workbook

Implants a dataframe into a specified cell range into an openxlsx2
workbook after validation checks have been completed.

## Usage

``` r
implant_df(
  wb,
  new_dat,
  cell_range,
  sheet,
  numeric_flag = TRUE,
  debug_mode = FALSE
)
```

## Arguments

- wb:

  An openxlsx2 workbook object

- new_dat:

  A dataframe containing the new data to be implanted

- cell_range:

  Character string specifying the range of cells to copy into (e.g.,
  "A1:C3")

- sheet:

  Character string or numeric index specifying the sheet to copy into

- numeric_flag:

  Logical indicating whether the data should be converted to numeric
  (default: TRUE)

- debug_mode:

  Logical indicating whether to highlight updated cells in red (default:
  FALSE)

## Value

Modified workbook object with data implanted

## Details

When `numeric_flag` is TRUE, all columns are converted to numeric. When
`debug_mode` is TRUE, updated cells are highlighted in blood red
(#880808) and an alert message is printed.
