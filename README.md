
# SKconvienience

<!-- badges: start -->
[![R-CMD-check](https://github.com/follhim/SKconvienience/actions/workflows/R-CMD-check.yaml/badge.svg)](https://github.com/follhim/SKconvienience/actions/workflows/R-CMD-check.yaml)
<!-- badges: end -->

SKconvienience is a package with convience functions for all things Rstats.

## Installation

You can install the development version of SKconvienience like so:

``` r
# Install devtools if you don't have it
install.packages("devtools")

# Install SKconvienience
devtools::install_github("follhim/SKconvienience")
```

## Example

This is a basic example which shows you how to solve a common problem:

```{r example, eval=FALSE}
library(SKconvienience)

# Basic correlation table
apa_cor_table(mtcars[, 1:5])

# With shifted descriptives (descriptive stats as rows at bottom)
apa_cor_table(mtcars[, 1:5], shift.descriptives = TRUE)

# Export to Excel
apa_cor_table(mtcars[, 1:5], filename = "my_correlation_table.xlsx")
```

