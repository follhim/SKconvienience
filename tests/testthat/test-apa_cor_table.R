test_that("apa_cor_table works without saving file", {
  # Create simple test data
  test_data <- data.frame(
    var1 = rnorm(30),
    var2 = rnorm(30),
    var3 = rnorm(30)
  )

  result <- apa_cor_table(test_data)
  expect_s3_class(result, "apa.cor.table")
})

test_that("apa_cor_table has correct structure", {
  test_data <- data.frame(
    var1 = rnorm(20),
    var2 = rnorm(20)
  )

  result <- apa_cor_table(test_data)

  # Check that it has the expected components
  expect_named(result, c("table.title", "table.body", "table.note", "landscape"))

  # Check that table.body is a matrix
  expect_true(is.matrix(result$table.body))
})

test_that("apa_cor_table handles shift.descriptives parameter", {
  test_data <- data.frame(
    var1 = rnorm(25),
    var2 = rnorm(25),
    var3 = rnorm(25)
  )

  result_shifted <- apa_cor_table(test_data, shift.descriptives = TRUE)

  expect_s3_class(result_shifted, "apa.cor.table")

  # With shifted descriptives, matrix should have extra rows for M, SD, N
  expect_true(nrow(result_shifted$table.body) > 3)
})

test_that("apa_cor_table removes non-numeric columns", {
  test_data <- data.frame(
    numeric1 = rnorm(10),
    numeric2 = rnorm(10),
    character = letters[1:10]
  )

  expect_warning(
    apa_cor_table(test_data),
    "Non-numeric columns have been removed"
  )
})
