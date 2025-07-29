##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## R Cervical Screening - KPIs for Publication
## func_cervical_opendata.R
## 
## Adult Screening Team
## Created: 23/08/2024
## Last update: 09/07/2025
## 
## Written/run on RStudio Posit Workbench
## Platform: x86_64-pc-linux-gnu
## R version 4.4.2 (Pile of Leaves) - 2024-10-31
##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


# Functions to add QF columns to OpenData dataframes ----

## OpenData HBRQF column ----
## df: dataframe containing a coulmn named HBR with hb2019 codes
od_hbrqf <- function(df) {
  df |>
    mutate(HBRQF = ifelse(HBR == "S92000003", "d", NA), 
           .after = HBR) }

## OpenData AgeGroupQF column ----
## df: dataframe containing a coulmn named AgeGroup
od_ageqf <- function(df) {
  df |>
    mutate(AgeGroupQF = ifelse(AgeGroup %in% c("25-49", "50-64", "25-64"), "d", NA), 
           .after = AgeGroup)
}

## OpenData SIMDQF column ----
## df: dataframe containing a coulmn names SIMD
od_simdqf <- function(df) {
  df |> 
    mutate(SIMDQF = ifelse(SIMD == "all", "d", NA),
           .after = SIMD)
}


