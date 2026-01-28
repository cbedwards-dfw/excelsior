# Changelog

## excelsior 0.0.1

- [`row_finder()`](https://cbedwards-dfw.github.io/excelsior/reference/row_finder.md):
  finds row number based on an “anchor” cell of known column and
  pattern.
- [`as_numeric_smart()`](https://cbedwards-dfw.github.io/excelsior/reference/as_numeric_smart.md):
  as.numeric(), but correctly handles cases of numbers with commas in
  them, numbers ending in %s.
- [`copy_columns()`](https://cbedwards-dfw.github.io/excelsior/reference/copy_columns.md):
  copies specified sets of columns from start_row to end_row from one
  file into another, with highlighting of cells copied and cells with
  changed values
- [`implant_df()`](https://cbedwards-dfw.github.io/excelsior/reference/implant_df.md)
  and
  [`copy_section()`](https://cbedwards-dfw.github.io/excelsior/reference/copy_section.md)
  to copy sections of excel sheets
- addition of
  [`clip_to_vec()`](https://cbedwards-dfw.github.io/excelsior/reference/clip_to_vec.md)
  to make coding easier
- addition of `validate_*()` for input validation

TODO: - confirm fram r-universe is including this
