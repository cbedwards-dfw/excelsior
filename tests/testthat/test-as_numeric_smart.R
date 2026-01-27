test_that("numeric transformation works", {
  res = as_numeric_smart(c("1,000", "10%", "5.5", "text"))
  expect_equal(res, c(1000, 0.1, 5.5, NA))
})
