# SKconvienience

<!-- badges: start -->
[![R-CMD-check](https://github.com/follhim/SKconvienience/actions/workflows/R-CMD-check.yaml/badge.svg)](https://github.com/follhim/SKconvienience/actions/workflows/R-CMD-check.yaml)
<!-- badges: end -->

SKconvienience is a package with convenience functions for all things R stats.

## Installation

You can install the development version of SKconvienience from GitHub:
``` r
# Install devtools if you don't have it
install.packages("devtools")

# Install SKconvienience
devtools::install_github("follhim/SKconvienience")
```

## Example

### Basic correlation table
``` r
library(SKconvienience)

apa_cor_table(mtcars[, 1:5])
```

### With shifted descriptives

Descriptive statistics displayed as rows at the bottom:
``` r
apa_cor_table(mtcars[, 1:5], shift.descriptives = TRUE)
```

### Export to Excel
``` r
apa_cor_table(mtcars[, 1:5], filename = "my_correlation_table.xlsx")
```

## Citation

To cite this package in publications:
``` r
citation("SKconvienience")
```
