test_that("validate_excel_column works", {
  foo <- function(m, n = NULL){
    validate_excel_column(m, n = n)
  }
  expect_no_error(foo("A"))
  expect_no_error(foo(c("A", "BB")))
  expect_error(foo(1))
  expect_error(foo("car"))

  expect_error(foo("A", n = 2))
  expect_no_error(foo(c("A", "B"), n = 2))

})


test_that("validate_hex_colors works", {
  foo <- function(m, n = NULL){
    validate_hex_color(m, n = n)
  }

  expect_no_error(foo("#FF0000"))

  expect_error(foo(c("red", "FF00", "#FF0000")))

  expect_error(foo(1))

  expect_error(foo(c("#82DE4E"), n = 2))

  expect_no_error(foo(c("#82DE4E", "#D5F4C3"), n = 2))
})


test_that("validate_cell_address works", {

  expect_no_error(validate_cell_address(c("A10", "B30", "AD4")))

  expect_error(validate_cell_address(1:2))

  expect_error(validate_cell_address(c("A10", "B30", "AD4"), n = 2))

  expect_error(validate_cell_address(c("A10", "B30", "AD4:34")))
})

test_that("validate_cell_range works", {

  expect_error(validate_cell_range(1:10))

  expect_no_error(validate_cell_range("A1:B10"))

  expect_no_error(validate_cell_range("A1"))

  expect_error(validate_cell_range("A1", n = 2))

  expect_error(validate_cell_range("A1", single_cell_allowed = FALSE))
})
