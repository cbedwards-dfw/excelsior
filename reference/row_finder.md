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

  Character or numeric. Column that contains anchor cell as number
  (e.g., 3) or corresponding excel column label ("C").

- pattern:

  Character. Pattern to identify anchor cell. Accepts regular
  expressions – see details.

- instance:

  Numeric. Which match to use if multiple matches found? Defaults to 1.

- offset:

  Numeric. Number to add to returned row number. Useful when pattern
  identifies a title row and you want the row below (1) or above (-1).
  Defaults to 0

## Value

Numeric. Row number where pattern was found, adjusted by offset

## Details

The `pattern` argument accepts regular expressions, which are more
flexible than just exactly matching a string. A few tips:

- R treats `\` specially – if present in the pattern, surround with
  square brackets to treat them as actual characters.

- In regular expressions, the `^` symbol is used to denote the start of
  a string, and the `$` symbol is used to denote the end of a string.

As an example of this in practice, if we had an anchor cell that was the
footnote "a/ Derived from vessel registrations and fish landing
tickets.", we could use as a pattern any section of the text (e.g.,
`pattern = "Derived from vessel registrations and fish"`). But if we
wanted to ensure that the first occurrence of "a/..." was used as our
anchor (in case the footnote text changes), we could use the pattern
`"^a[/]"`, which in words means "the first letter is an a, and the
second letter is literally a `/`".

## Examples

``` r
if (FALSE) { # \dontrun{
row_finder("data.xlsx", "Sheet1", "A", "Total", instance = 1, offset = 1)
row_finder("data.xlsx", 1, 3, "Summary", offset = -1)
} # }
```
