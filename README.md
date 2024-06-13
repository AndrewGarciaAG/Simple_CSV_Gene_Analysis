# Gene Matching and Analysis Script

This repository contains an R script for processing Excel sheets to identify matching genes, perform comparisons, and analyze specific columns. The script reads data from multiple sheets, finds matching genes, writes the results to CSV files, and calculates the range for specified columns.

## Features

- Reads multiple sheets from an Excel file.
- Identifies and prints matching genes.
- Finds common matches across selected sheets.
- Calculates the range (min and max) for specific columns.
- Creates tables for matched genes and their scores.
- Writes results to CSV files.

## Requirements

- R (version 4.0.0 or higher)
- `readxl` package

## Installation

1. Install R from [CRAN](https://cran.r-project.org/).
2. Install the `readxl` package if not already installed:

```r
install.packages("readxl")
