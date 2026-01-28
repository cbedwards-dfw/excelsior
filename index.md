# excelsior

The goal of excelsior is to provide tools for automating data entry that
involves transferring data from one excel file to another.

## Installation

You can install the development version of excelsior from
[GitHub](https://github.com/) with:

``` r
# install.packages("pak")
pak::pak("cbedwards-dfw/excelsior")
```

If you do not have `Rtools` installed, you can install this package (and
`xldiff`, upon which it depends) from the FRAMverse R universe:

``` r
install.packages(c('excelsior', 'xldiff'), repos = c('https://framverse.r-universe.dev', 'https://cloud.r-project.org'))
```
