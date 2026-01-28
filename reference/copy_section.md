# Copy data section between locations in Excel workbook

Designed with updating in mind. Copies data from one location to another
within an Excel workbook, with optional validation checks for matching
reference rows/columns. This is a safe version of saying "Copy sheet
C3:D4 of sheet 1 to E5:F6 of sheet 2". To support speed, this function
takes as arguments the openxlsx2 workbook object to be updated, and a
dataframe to update FROM.

## Usage

``` r
copy_section(
  wb,
  from_df,
  from_address,
  to_address,
  sheet,
  check_row_offset = NULL,
  check_col_offset = NULL,
  numeric_flag = TRUE,
  debug_mode = FALSE
)
```

## Arguments

- wb:

  An openxlsx2 workbook object

- from_df:

  A dataframe of the complete sheet that contains the data to copy
  *from*. Typically generated from
  [`readxl::read_excel`](https://readxl.tidyverse.org/reference/read_excel.html)
  with `col_names = FALSE` and `.name_repair = "unique_quiet`.

- from_address:

  Character string specifying the source cell range (e.g., "A1:C3").

- to_address:

  Character string specifying the destination cell range (e.g., "D1:F3")

- sheet:

  Character string or numeric index specifying the target sheet

- check_row_offset:

  Optional, integer offset identifying the row that should be used to
  validate that the from/to sections match. Positive values correspond
  to rows *above* the first row of the addresses specified.

- check_col_offset:

  Optional, integer offset identifying the column that should be used to
  validate that the from/to sections match. Positive values correspond
  to columns *before* the first column of the addresses specified.

- numeric_flag:

  Logical indicating whether data should be converted to numeric
  (default: TRUE)

- debug_mode:

  Logical indicating whether to enable debug mode highlighting (default:
  FALSE)

## Value

Modified workbook object with data copied

## Details

The key feature of this function is optional checking, in which the
first column or row of the "from_db" is checked against a corresponding
column or row of the "to_address" to make sure the entries are still
appropriate.

As an example, when updating the TAMM template file with a new year's
data, we have a rectangular section of locations and their values
present in the TAMM and in the input template; we want to update the
input template with the new TAMM values. We can code this with cell
addresses, identifying that we want to copy from cells "D126:F130" of
the TAMM to "D5:F9" of the current sheet of the input template. However,
we could encounter problems if users have re-arranged either excel
sheet. For this reason, we want to confirm that the location column
(which is 1 column to the left of the leftmost cells we're copying)
matches. So we provide "check_col_offset = 1" as an argument. This
causes the function to check that "C123:C130" of the TAMM matches
"C5:C9" of the current sheet of the input template.
