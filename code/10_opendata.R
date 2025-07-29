##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## R Cervical Screening - KPIs for Publication
## 10_opendata.R
## 
## Adult Screening Team
## Created: 26/05/25
## Last update: 09/07/2025
## 
## Written/run on RStudio Posit Workbench
## Platform: x86_64-pc-linux-gnu
## R version 4.4.2 (Pile of Leaves) - 2024-10-31
##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Packages ----
library(dplyr)
library(readr)

## Clean global environment
rm(list=ls())


## Sourcing
source("code/1_housekeeping.R")
source("src/func_cervical_opendata.R")


# Import rds files ----
## KPI 1 ----
kpi1_1 <- readRDS(paste0(fy_folder_path,
                         "historic/kpi1_1_hist_coverage_", year_label, ".rds"))
kpi1_1_simd <- readRDS(paste0(fy_folder_path,
                              "historic/kpi1_1_hist_coverage_simd_", year_label, ".rds"))
kpi1_2 <- readRDS(paste0(fy_folder_path,
                         "historic/kpi1_2_hist_uptake_", year_label, ".rds"))
kpi1_2_simd <- readRDS(paste0(fy_folder_path,
                              "historic/kpi1_2_hist_uptake_simd_", year_label, ".rds"))

## KPI 2 ----
kpi2_1 <- read_rds(paste0(fy_folder_path,
                          "historic/kpi2_1_hist_reporting_", year_label, ".rds"))
kpi2_2 <- read_rds(paste0(fy_folder_path,
                          "historic/kpi2_2_hist_rejected_", year_label, ".rds"))

## KPI 3 ----
kpi3_1 <- read_rds(paste0(fy_folder_path,
                          "historic/kpi3_1_hist_cytology_unsatis_", year_label, ".rds"))
kpi3_2 <- read_rds(paste0(fy_folder_path,
                          "historic/kpi3_2_hist_hpv_failed_", year_label, ".rds"))
kpi3_3 <- read_rds(paste0(fy_folder_path,
                          "historic/kpi3_3_hist_repeat_test_", year_label, ".rds"))


# KPI 1 Open Data preparation ----
od_kpi1_1 <- kpi1_1 |> 
  select(!hbres_screen)|>
  rename(FinancialYear = fin_year,
         HBR = hb2019,
         AgeGroup = age_group,
         NumberEligible = eligible,
         NumberAgeAppropriateScreen = screened,
         PercentageCoverage = age_appr_coverage) |> # CamelCase
  mutate(SIMD = "all")

od_kpi1_1_simd <- kpi1_1_simd |> 
  select(!hbres_screen)|>
  rename(FinancialYear = fin_year,
         HBR = hb2019,
         SIMD = simd,
         AgeGroup = age_group,
         NumberEligible = eligible,
         NumberAgeAppropriateScreen = screened,
         PercentageCoverage = age_appr_coverage) |> # CamelCase
  rbind(od_kpi1_1) |>
  od_hbrqf() |> 
  od_ageqf() |> 
  od_simdqf() |> 
  mutate(NumberEligible = ifelse(is.na(NumberEligible), 0, NumberEligible),
         NumberAgeAppropriateScreen = ifelse(is.na(NumberAgeAppropriateScreen), 0, NumberAgeAppropriateScreen),
         PercentageCoverageQF = ifelse(is.na(PercentageCoverage), "z", NA)) |>
  arrange(desc(FinancialYear), HBR, AgeGroup) 

od_kpi1_2 <- kpi1_2 |> 
  select(!hbres_invite)|>
  rename(FinancialYear = fin_year,
         HBR = hb2019,
         AgeGroup = age_group,
         NumberInvite = invites,
         NumberScreen = screened,
         PercentageUptake = uptake) |> # CamelCase
  mutate(SIMD = "all")

od_kpi1_2_simd <- kpi1_2_simd |> 
  select(!hbres_invite)|>
  rename(FinancialYear = fin_year,
         HBR = hb2019,
         SIMD = simd,
         AgeGroup = age_group,
         NumberInvite = invites,
         NumberScreen = screened,
         PercentageUptake = uptake) |> # CamelCase
  rbind(od_kpi1_2) |>
  od_hbrqf() |> 
  od_ageqf() |> 
  od_simdqf() |> 
  mutate(NumberInvite = ifelse(is.na(NumberInvite), 0, NumberInvite),
         NumberScreen = ifelse(is.na(NumberScreen), 0, NumberScreen),
         PercentageUptakeQF = ifelse(is.na(PercentageUptake), "z", NA)) |>
  arrange(desc(FinancialYear), HBR, AgeGroup)


# KPI 2 Open Data preparation ----
od_kpi2_1 <- kpi2_1 |> 
  select(!hbres_screen) |>
  rename(FinancialYear = fin_year,
         HBR = hb2019,
         NumberOfResultsReported = total_reports,
         NumberOfReportsWithin2Weeks = report_2week,
         PercentageOfReportsWithin2Weeks = report_2week_pct) |>
  od_hbrqf()

od_kpi2_2 <- kpi2_2 |> 
  select(!hbres_screen) |>
  rename(FinancialYear = fin_year,
         HBR = hb2019,
         NumberOfSamplesReceived = total_samples,
         NumberOfSamplesRejected = rejected_samples,
         PercentageOfSamplesRejected = rejected_samples_pct) |>
  od_hbrqf()


# KPI 3 Open Data preparation ----
od_kpi3_1 <- kpi3_1 |> 
  select(!hbres_screen) |>
  rename(FinancialYear = fin_year,
         HBR = hb2019,
         NumberOfCytologySamples = cytology_samples,
         NumberOfUnsatisfactoryCytologyResults = unsatisfactory,
         PercentageOfUnsatisfactoryCytologyResults = unsatisfactory_pct) |>
  od_hbrqf()

od_kpi3_2 <- kpi3_2 |> 
  select(!hbres_screen) |>
  rename(FinancialYear = fin_year,
         HBR = hb2019,
         NumberOfVirologySamples = virology_samples,
         NumberOfHPVFailResults = hpv_fail,
         PercentageOfHPVFailResults = hpv_fail_pct) |>
  od_hbrqf()

od_kpi3_3 <- kpi3_3 |> 
  select(!hbres_screen) |>
  rename(FinancialYear = fin_year,
         HBR = hb2019,
         NumberOfFailedResults = result_fail,
         NumberOfRepeatScreen = repeat_screen,
         PercentageOfRepeatScreen = repeat_screen_pct) |>
  od_hbrqf()


# Save OpenData csv files ----
write_csv(od_kpi1_1_simd,
          file = paste0(fy_folder_path, "opendata/",
                        "Cervical_Screening_Coverage_201617_20", year_label, ".csv"),
          na = "")
write_csv(od_kpi1_2_simd,
          file = paste0(fy_folder_path, "opendata/",
                        "Cervical_Screening_Uptake_201617_20", year_label, ".csv"),
          na = "")

write_csv(od_kpi2_1,
          file = paste0(fy_folder_path, "opendata/",
                        "Cervical_Screening_LabResults_201617_20", year_label, ".csv"),
          na = "")
write_csv(od_kpi2_2,
          file = paste0(fy_folder_path, "opendata/",
                        "Cervical_Screening_LabResultsRejected_201617_20", year_label, ".csv"),
          na = "")

write_csv(od_kpi3_1,
          file = paste0(fy_folder_path, "opendata/",
                        "Cervical_Screening_CytologySamples_201617_20", year_label, ".csv"),
          na = "")
write_csv(od_kpi3_2,
          file = paste0(fy_folder_path, "opendata/",
                        "Cervical_Screening_VirologySamples_201617_20", year_label, ".csv"),
          na = "")
write_csv(od_kpi3_3,
          file = paste0(fy_folder_path, "opendata/",
                        "Cervical_Screening_RepeatScreen_201617_20", year_label, ".csv"),
          na = "")


# End of Script ----
