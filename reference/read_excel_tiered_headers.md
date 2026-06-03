# Read an Excel file with multi-row headers

Reads an Excel worksheet where column headers are spread across multiple
rows, including merged cells. Header rows are combined into a single
header string per column, with consecutive duplicate values collapsed
(e.g. a superheader "TREATY" spanning columns that are individually
labelled "TROLL", "NET", "SPORT" becomes "TREATY_TROLL", "TREATY_NET",
"TREATY_SPORT").

## Usage

``` r
read_excel_tiered_headers(
  path,
  sheet = 1,
  header_rows,
  pseudo_merged_rows = NULL,
  first_data_row = NULL,
  final_data_row = NULL,
  first_column = NULL,
  final_column = NULL,
  sep = "_",
  clean_names = TRUE
)
```

## Arguments

- path:

  Character. Path to the `.xlsx` file.

- sheet:

  Integer or character atomic Sheet index to read. Default `1` to
  automatically handle cases with a single sheet.

- header_rows:

  Integer vector. Row numbers (1-indexed, as they appear in the
  spreadsheet) that together form the column headers.

- pseudo_merged_rows:

  Integer vector. Row numbers (1-indexed, as they appear in the
  spreadsheet) for header rows that include "false merged cells" – cells
  that look like they're merged because the first cell of a group has
  weird formatting that shifts the text far to the right. Warning: this
  can create funky names cases in which there are gaps columns in the
  merged row. This is NOT necessary if the excel cells are *actually*
  merged. Defaults to NULL.

- first_data_row:

  Integer or `NULL`. Row number of the first data row. If `NULL`
  (default), set to `max(header_rows) + 1`.

- final_data_row:

  Integer or `NULL`. Row number of the last data row. If `NULL`
  (default), all rows from `first_data_row` to the end of the sheet are
  returned.

- first_column:

  Integer or `NULL`. Column number of the first column to read in. If
  `NULL` (default), uses default behavior of
  [`openxlsx2::read_xlsx()`](https://janmarvin.github.io/openxlsx2/reference/wb_to_df.html).

- final_column:

  Integer or `NULL`. Column number of the last column to read in. If
  `NULL` (default), uses default behavior of
  [`openxlsx2::read_xlsx()`](https://janmarvin.github.io/openxlsx2/reference/wb_to_df.html).

- sep:

  Character. Separator used when joining header parts from different
  rows. Default `"_"`.

- clean_names:

  Logical. Resolves issues in which column names are not unique. If
  `TRUE` (default), column names are passed through
  [`base::make.names()`](https://rdrr.io/r/base/make.names.html) with
  `unique = TRUE` to ensure they are valid and non-duplicate R names.

## Value

A data frame with one column per spreadsheet column and one row per data
row. Column names are the combined headers derived from `header_rows`.

## Details

Merged cells are handled via
[`openxlsx2::read_xlsx()`](https://janmarvin.github.io/openxlsx2/reference/wb_to_df.html)
with `fill_merged_cells = TRUE`, which propagates the anchor cell's
value across the full span of the merge before any header combination is
attempted.

## Examples
