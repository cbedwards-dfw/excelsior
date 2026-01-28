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
