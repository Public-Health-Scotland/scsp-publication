##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## R Cervical Screening - KPIs for Publication
## 4_cervical_kpi1_2549_5064.R
## 
## Adult Screening Team
## Created: 26/07/24
## Last update: 09/07/2025
## 
## Written/run on RStudio Posit Workbench
## Platform: x86_64-pc-linux-gnu
## R version 4.4.2 (Pile of Leaves) - 2024-10-31
##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Packages ----
library(readr)
library(dplyr)
library(tidyr)
library(tidylog)
library(lubridate)
library(forcats)
library(phsmethods)
library(ggplot2)
library(openxlsx)

## Clean global environment
rm(list=ls())


## Sourcing
source("code/1_housekeeping.R")
source("src/func_cervical_kpi_calc_graph.R")
source("src/func_cervical_wrt_kpi_data.R")
source("src/func_cervical_wrt_xlsx_workbook.R")
source("src/func_cervical_wrt_xlsx_workbook_2549_5064.R")


## Clean global environment
rm(list=ls(pattern = "kpi[2-3]"))

# Variables ----
scr_period_5yr <- paste0("The ", fin_yr, " 5.5-year look-back period: 1st October ", 
                                year(dmy(fy_start_date_5_5)), "- 31st March ", yr_end)
scr_period_3yr <- paste0("The ", fin_yr, " 3.5-year look-back period: 1st October ", 
                                year(dmy(fy_start_date_3_5)), " - 31st March ", yr_end)

kpi1_title <- "KPI 1: Coverage, 3.5 year and 5.5 year split"

kpi_notes <- list(
  kpi1_notes = list(
    content = tibble::tibble(
      content = c("Key Performance Indicator (KPI) Report",
                  paste("Reporting period:", scr_period_1yr),
                  "",
                  "  1. Coverage and Uptake",
                  paste("      KPI 1.1  Coverage of eligible women aged 25-49 who have attended", 
                        "cervical screening within 3.5 years"),
                  paste("      KPI 1.1  Coverage of eligible women aged 50-64 who have attended", 
                        "cervical screening within 5.5 years"),
                  paste("      KPI 1.1a Coverage of eligible women aged 25-49 who have attended",
                        "screening within 3.5 years by deprivation (SIMD)"),
                  paste("      KPI 1.1a Coverage of eligible women aged 50-64 who have attended",
                        "screening within 5.5 years by deprivation (SIMD)")
      )),
    def = tibble::tibble(
      def = c("Definitions:",
              paste("Coverage is defined as the percentage",
                    "of women in a population eligible for",
                    "screening at a given point in time who",
                    "were screened adequately within a specific period.")
      )),
    notes = tibble::tibble(
      notes = c("Notes:",
                paste("Whilst this workbook uses the term ‘women’,",
                      "not only those who identify as women need",
                      "access to cervical screening. Transgender men,",
                      "non-binary and intersex people with a cervix are eligible."),
                "From 2016/17 the eligible age for cervical screening is 25-64.",
                paste0("Historically, women aged 25-49 have been invited ",
                       "to attend cervical screening every three years ",
                       "and those aged 50-64 every five. From March 2020, ",
                       "women aged 25-64 are invited every five years. ",
                       "This will be implemented to the KPI analysis in ",
                       all_scr5_5_start, "/", all_scr5_5_start+1,
                       ", to allow screening participant to move to the new recall."),
                "",
                "            Table 1: Age-appropriate look-back period at a given point in time",
                "", #kpi_notes$kpi1_cov_periods_table,
                "",
                "",
                "",
                "",
                ""
      )),
    table = tibble::tibble(
      col_a = c("Reporting period\n(01 April - 31 March)",
                fin_yr,
                paste0(yr_start-1, "/", as.double(substr(yr_end, 3, 4))-1),
                paste0(yr_start-2, "/", as.double(substr(yr_end, 3, 4))-2),
                paste0(yr_start-3, "/", as.double(substr(yr_end, 3, 4))-3),
                paste0(yr_start-4, "/", as.double(substr(yr_end, 3, 4))-4)),
      col_b = c("Age group",
                ifelse(yr_end>all_scr5_5_start, "25-64", "25-49\n50-64"),
                ifelse(yr_end-1>all_scr5_5_start, "25-64", "25-49\n50-64"),
                ifelse(yr_end-2>all_scr5_5_start, "25-64", "25-49\n50-64"),
                ifelse(yr_end-3>all_scr5_5_start, "25-64", "25-49\n50-64"),
                ifelse(yr_end-4>all_scr5_5_start, "25-64", "25-49\n50-64")),
      col_c = c("Length of look-back",
                ifelse(yr_end>all_scr5_5_start, "5.5 year", "3.5 year\n5.5 year"),
                ifelse(yr_end-1>all_scr5_5_start, "5.5 year", "3.5 year\n5.5 year"),
                ifelse(yr_end-2>all_scr5_5_start, "5.5 year", "3.5 year\n5.5 year"),
                ifelse(yr_end-3>all_scr5_5_start, "5.5 year", "3.5 year\n5.5 year"),
                ifelse(yr_end-4>all_scr5_5_start, "5.5 year", "3.5 year\n5.5 year")),
      col_d = c("Look-back period",
                ifelse(yr_end>all_scr5_5_start, 
                       paste("01 October", yr_start-5, "to 31 March", yr_end), 
                       paste("01 October", yr_start-3, "to 31 March", yr_end,
                             "\n01 October", yr_start-5, "to 31 March", yr_end)),
                ifelse(yr_end-1>all_scr5_5_start,
                       paste("01 October", yr_start-6, "to 31 March", yr_end-1), 
                       paste("01 October", yr_start-4, "to 31 March", yr_end-1,
                             "\n01 October", yr_start-6, "to 31 March", yr_end-1)),
                ifelse(yr_end-2>all_scr5_5_start,
                       paste("01 October", yr_start-7, "to 31 March", yr_end-2), 
                       paste("01 October", yr_start-5, "to 31 March", yr_end-2,
                             "\n01 October", yr_start-7, "to 31 March", yr_end-2)),
                ifelse(yr_end-3>all_scr5_5_start,
                       paste("01 October", yr_start-8, "to 31 March", yr_end-3), 
                       paste("01 October", yr_start-6, "to 31 March", yr_end-3,
                             "\n01 October", yr_start-8, "to 31 March", yr_end-3)),
                ifelse(yr_end-4>all_scr5_5_start,
                       paste("01 October", yr_start-9, "to 31 March", yr_end-4), 
                       paste("01 October", yr_start-7, "to 31 March", yr_end-4,
                             "\n01 October", yr_start-9, "to 31 March", yr_end-4))
      )),
    source = tibble::tibble(
      source = c("Source:",
                 "Based on the Scottish Cervical Call Recall System (SCCRS).",
                 paste0("Data download date: ", substr(extract_dl_date, 7, 8), 
                        "-", substr(extract_dl_date, 5, 6), "-",
                        substr(extract_dl_date, 1, 4))
      ))
  )
)

kpi1_list <- list(kpi1_title = "KPI 1 - Coverage, 3.5 year and 5.5 year split", 
                  kpi1_1_subt_35 = paste("KPI 1.1: Coverage of eligible women aged 25-49 who have attended", 
                                      "cervical screening within 3.5 years"),
                  kpi1_1_subt_55 = paste("KPI 1.1: Coverage of eligible women aged 50-64 who have attended", 
                                         "cervical screening within 5.5 years"),
                  fig1_1_legend_35 = paste("Figure 1.1: Coverage (%) of eligible women with a", 
                                        "screening encounter within the previous 3.5 year", 
                                        "by NHS Health Board of Residence,", fin_year_5[1], "to", fin_yr),
                  fig1_1_legend_55 = paste("Figure 1.1: Coverage (%) of eligible women with a", 
                                           "screening encounter within the previous 5.5 year", 
                                           "by NHS Health Board of Residence,", fin_year_5[1], "to", fin_yr),
                  table1_1_header_35 = paste("Table 1.1: Coverage (%) of eligible women screened within 3.5 years", 
                                          "by NHS Health Board of Residence,", first_report_fin_yr, "to", fin_yr),
                  table1_1_header_55 = paste("Table 1.1: Coverage (%) of eligible women screened within 5.5 years", 
                                             "by NHS Health Board of Residence,", first_report_fin_yr, "to", fin_yr),
                  kpi1_1a_subt_35 = paste("KPI 1.1a: Coverage of eligible women aged 25-49 who have attended", 
                                       "screening within 3.5 years by deprivation (SIMD)"),
                  kpi1_1a_subt_55 = paste("KPI 1.1a: Coverage of eligible women aged 50-64 who have attended", 
                                          "screening within 5.5 years by deprivation (SIMD)"),
                  fig1_1a_legend_35 = paste("Figure 1.1a: Scotland Coverage (%) of eligible women with a", 
                                         "screening encounter within the previous 3.5 years", 
                                         "by deprivation (SIMD),", fin_year_5[1], "to", fin_yr),
                  fig1_1a_legend_55 = paste("Figure 1.1a: Scotland Coverage (%) of eligible women with a", 
                                            "screening encounter within the previous 5.5 years", 
                                            "by deprivation (SIMD),", fin_year_5[1], "to", fin_yr),
                  table1_1a_hb_header_35 = paste("Table 1.1a1: NHS Health Board of Residence and Scotland", 
                                              "Coverage (%) within 3.5 years by deprivation (SIMD), financial year", fin_yr),
                  table1_1a_hb_header_55 = paste("Table 1.1a1: NHS Health Board of Residence and Scotland", 
                                                 "Coverage (%) within 5.5 years by deprivation (SIMD), financial year", fin_yr),
                  table1_1a_sc_header_35 = paste("Table 1.1a2: Scotland Coverage (%) within 3.5 years by deprivation (SIMD),",
                                                 "financial years,", first_report_fin_yr, "to", fin_yr),
                  table1_1a_sc_header_55 = paste("Table 1.1a2: Scotland Coverage (%) within 5.5 years by deprivation (SIMD),",
                                                 "financial years,", first_report_fin_yr, "to", fin_yr))

save_to_kpi1_sd_xlsx <- paste0(fy_folder_path,
                               "output/KPI_1_Supplementary_Data_", year_label, ".xlsx")



# Import KPI 1 rds files ----
hist_cov_add <- read_rds(paste0(fy_folder_path,
                                "historic/kpi1_1_hist_coverage_", year_label, ".rds"))

hist_simd_cov_add <- read_rds(paste0(fy_folder_path,
                                     "historic/kpi1_1_hist_coverage_simd_", year_label, ".rds"))


# Prepare data for publication ----

# KPI 1 - Coverage by HBs, 5 year trend data ----
## 25-49 ----
hist_cov_table_2549 <- hist_cov_add %>% 
  filter(age_group == "25-49") %>% 
  select(!c(hb2019, age_group, eligible, screened)) 

### Table for printing
kpi1_1_5yr_2549 <- hist_cov_table_2549 %>% 
  #filter(fin_year %in% fin_year_5) %>% 
  pivot_wider(names_from = fin_year,
              values_from = age_appr_coverage) %>% 
  rename("Health Board of Residence" = hbres_screen)

### Graph for printing
kpi1_1_graph_2549 <- cervical_column_graph(df = hist_cov_table_2549 %>%
                                        filter(fin_year %in% fin_year_5) %>%
                                        select("Health Board of Residence" = hbres_screen,
                                               "Financial Year" = fin_year, 
                                               "Coverage (%)\nWomen aged 25-49" = age_appr_coverage),
                                      data_col = "Coverage (%)\nWomen aged 25-49", 
                                      colours = create_palette(hist_cov_table_2549 %>%
                                                                 filter(fin_year %in% fin_year_5) |> 
                                                                 pull(fin_year)), 
                                      threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/supp_kpi1_1_2549_", year_label, ".png"),
       plot = kpi1_1_graph_2549,
       width = 9, height = 5)

## 50-64 ----
hist_cov_table_5064 <- hist_cov_add %>% 
  filter(age_group == "50-64") %>% 
  select(!c(hb2019, age_group, eligible, screened)) 

### Table for printing
kpi1_1_5yr_5064 <- hist_cov_table_5064 %>% 
  #filter(fin_year %in% fin_year_5) %>% 
  pivot_wider(names_from = fin_year,
              values_from = age_appr_coverage) %>% 
  rename("Health Board of Residence" = hbres_screen)

### Graph for printing
kpi1_1_graph_5064 <- cervical_column_graph(df = hist_cov_table_5064 %>%
                                             filter(fin_year %in% fin_year_5) %>%
                                             select("Health Board of Residence" = hbres_screen,
                                                    "Financial Year" = fin_year, 
                                                    "Coverage (%)\nWomen aged 50-64" = age_appr_coverage),
                                           data_col = "Coverage (%)\nWomen aged 50-64", 
                                           colours = create_palette(hist_cov_table_5064 %>%
                                                                      filter(fin_year %in% fin_year_5) |> 
                                                                      pull(fin_year)), 
                                           threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/supp_kpi1_1_5064_", year_label, ".png"),
       plot = kpi1_1_graph_5064,
       width = 9, height = 5)

# KPI 1 - Coverage by SIMD, 5 year trend data ----
## 25-49 ----
hist_simd_cov_table_2549 <- hist_simd_cov_add %>% 
  filter(age_group == "25-49") %>% 
  select(!c(hb2019, age_group, eligible, screened)) %>% 
  mutate(simd = recode(simd, "1" = "1 - Most deprived", "5" = "5 - Least deprived"))

# Tables for printing
kpi1_1a_scotland_2549 <- hist_simd_cov_table_2549 %>% 
  #filter(fin_year %in% fin_year_5) %>% 
  filter(hbres_screen == "Scotland") 

kpi1_1a_5yr_table_2549 <- kpi1_1a_scotland_2549 %>% 
  pivot_wider(names_from = fin_year,
              values_from = age_appr_coverage)  %>% 
  select(!hbres_screen) %>% 
  rename("Deprivation (SIMD)" = simd)

kpi1_1a_1yr_table_2549 <- hist_simd_cov_table_2549 %>% 
  filter(fin_year == fin_yr) %>% 
  select(!fin_year) %>% 
  pivot_wider(names_from = hbres_screen,
              values_from = age_appr_coverage) %>% 
  rename("Deprivation (SIMD)"= simd)

# Graph for printing
kpi1_1a_graph_2549 <- cervical_column_graph(df = kpi1_1a_scotland_2549 %>%
                                         filter(fin_year %in% fin_year_5) %>%
                                         select("Deprivation (SIMD)" = simd, "Financial Year" = fin_year,
                                                "Coverage (%)\nWomen aged 25-49" = age_appr_coverage),
                                      data_col = "Coverage (%)\nWomen aged 25-49", 
                                      colours = create_palette(kpi1_1a_scotland_2549 %>%
                                                                 filter(fin_year %in% fin_year_5) |> 
                                                                 pull(fin_year)),
                                      threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/supp_kpi1_1a_2549_", year_label, ".png"),
       plot = kpi1_1a_graph_2549,
       width = 7.5, height = 4)

## 50-64 ----
hist_simd_cov_table_5064 <- hist_simd_cov_add %>% 
  filter(age_group == "50-64") %>% 
  select(!c(hb2019, age_group, eligible, screened)) %>% 
  mutate(simd = recode(simd, "1" = "1 - Most deprived", "5" = "5 - Least deprived"))

# Tables for printing
kpi1_1a_scotland_5064 <- hist_simd_cov_table_5064 %>% 
  #filter(fin_year %in% fin_year_5) %>% 
  filter(hbres_screen == "Scotland") 

kpi1_1a_5yr_table_5064 <- kpi1_1a_scotland_5064 %>% 
  pivot_wider(names_from = fin_year,
              values_from = age_appr_coverage)  %>% 
  select(!hbres_screen) %>% 
  rename("Deprivation (SIMD)" = simd)

kpi1_1a_1yr_table_5064 <- hist_simd_cov_table_5064 %>% 
  filter(fin_year == fin_yr) %>% 
  select(!fin_year) %>% 
  pivot_wider(names_from = hbres_screen,
              values_from = age_appr_coverage) %>% 
  rename("Deprivation (SIMD)"= simd)

# Graph for printing
kpi1_1a_graph_5064 <- cervical_column_graph(df = kpi1_1a_scotland_5064 %>%
                                              filter(fin_year %in% fin_year_5) %>%
                                              select("Deprivation (SIMD)" = simd, "Financial Year" = fin_year,
                                                     "Coverage (%)\nWomen aged 50-64" = age_appr_coverage),
                                            data_col = "Coverage (%)\nWomen aged 50-64",
                                            colours = create_palette(kpi1_1a_scotland_5064 %>%
                                                                    filter(fin_year %in% fin_year_5) |> 
                                                                    pull(fin_year)),
                                            threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/supp_kpi1_1a_5064_", year_label, ".png"),
       plot = kpi1_1a_graph_5064,
       width = 7.5, height = 4)


# Write KPI 1 workbook ----
wrt_xlsx_kpi1_2549_5064(scr_period_1yr, scr_period_3yr, scr_period_5yr, fmt_styles, 
                        kpi_notes$kpi1_notes, kpi1_list,
                        kpi1_1_5yr_2549, kpi1_1_graph_2549,
                        kpi1_1a_1yr_table_2549, kpi1_1a_5yr_table_2549, kpi1_1a_graph_2549,
                        kpi1_1_5yr_5064, kpi1_1_graph_5064,
                        kpi1_1a_1yr_table_5064, kpi1_1a_5yr_table_5064, kpi1_1a_graph_5064,
                        save_to_kpi1_sd_xlsx)


# End of Script ----

