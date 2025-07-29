##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## R Cervical Screening - KPIs for Publication
## 5_cervical_kpi2.R
## 
## Adult Screening Team
## Created: 29/08/24
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
library(janitor)
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

## Clean global environment
rm(list=ls(pattern = "kpi1|kpi[3]|fig1|fig[3]|simd"))


# Import extract ----
# If KPI analysis has previously been prepared for the financial year of interest
# Jump to section: Prepare data for publication
kpi_extract <- read_rds(paste0(fy_folder_path,
                               "temp/kpi2_extract_", year_label, ".rds")) %>%
  # remove any data without health board information
  filter(!is.na(`NHS Board of Residence`))

# Import historic rds files ----
## KPI 2.1
hist_reporting <- readRDS(paste0(hist_folder_path, "/kpi2_1_hist_reporting_", 
                                 substr(yr_start-1, 3, 4), substr(yr_end-1, 3, 4),
                                 ".rds"))

## KPI 2.2
hist_rejected <- readRDS(paste0(hist_folder_path, "/kpi2_2_hist_rejected_", 
                                substr(yr_start-1, 3, 4), substr(yr_end-1, 3, 4),
                                ".rds"))


# KPI 2 - Laboratory Service KPIs ----
## KPI 2.1 - Annual percentage of samples reported within 10 working days ----
## from the date of the sample being taken
kpi2_1 <- cervical_pct(kpi_extract, (reported_fy == 1),
                       reported_14, reported_fy,
                       hb2019) %>% 
  rename("total_reports" = Denominator,
         "report_2week" = Numerator,
         "report_2week_pct" = Percentage)

## KPI 2.2 - Percentage of tests rejected by the laboratory prior to processing ----
kpi2_2 <- cervical_pct(kpi_extract, (arrived_fy == 1),
                       rejected, arrived_fy,
                       hb2019) %>% 
  rename("total_samples" = Denominator,
         "rejected_samples" = Numerator,
         "rejected_samples_pct" = Percentage)


# Save the new KPI 2 to historic records ----
hist_reporting_add <- rbind(hist_reporting, 
                            kpi2_1 %>%
                              mutate(fin_year = fin_yr, 
                                     hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                                                           levels = hbres_levels),
                                     .before = 1)) %>% 
  arrange(hbres_screen)

saveRDS(hist_reporting_add,
        file = paste0(fy_folder_path,
                      "historic/kpi2_1_hist_reporting_", year_label, ".rds"))

hist_rejected_add <- rbind(hist_rejected, 
                           kpi2_2 %>%
                             mutate(fin_year = fin_yr, 
                                    hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                                                          levels = hbres_levels),
                                    .before = 1)) %>% 
  arrange(hbres_screen)

saveRDS(hist_rejected_add, 
        file = paste0(fy_folder_path,
                      "historic/kpi2_2_hist_rejected_", year_label, ".rds"))


# Prepare data for publication ----
# If above KPI data has previously been prepared, import rds files:
# hist_reporting_add <- read_rds(paste0(fy_folder_path,
#                                       "historic/kpi2_1_hist_reporting_", year_label, ".rds"))
# hist_rejected_add <- read_rds(paste0(fy_folder_path,
#                                      "historic/kpi2_2_hist_rejected_", year_label, ".rds"))

## KPI 2.1 - Number of samples reported within 2-weeks ----
hist_report_table <- hist_reporting_add %>% 
  select(!c(hb2019, total_reports, report_2week))

# Table for printing
kpi2_1_table <- hist_report_table %>% 
  #filter(fin_year %in% fin_year_5) %>% # add back in after 2022/23 and 2023/24 combined publication 
  pivot_wider(names_from = fin_year,
              values_from = report_2week_pct) %>% 
  rename("Health Board of Residence" = hbres_screen,
         "2020/21\u00b9" = "2020/21")

# Graph for printing
kpi2_1_graph <- cervical_column_graph(df = hist_report_table %>%
                                        filter(fin_year %in% fin_year_5) %>%
                                        select("Health Board of Residence" = hbres_screen,
                                               "Financial Year" = fin_year, 
                                               "Reported within 2-weeks (%)" = report_2week_pct),
                                      data_col = "Reported within 2-weeks (%)", 
                                      colours = create_palette(hist_report_table  |> 
                                                                 filter(fin_year %in% fin_year_5) |> 
                                                                 pull(fin_year)),
                                      threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi2_1_", year_label, ".png"),
       plot = kpi2_1_graph,
       width = 9, height = 5)

## KPI 2.2 - Number of samples rejected by lab ----
hist_rejected_table <- hist_rejected_add %>% 
  select(!c(hb2019, total_samples, rejected_samples))

# Table for printing
kpi2_2_table <- hist_rejected_table %>% 
  #filter(fin_year %in% fin_year_5) %>% # add back in after 2022/23 and 2023/24 combined publication 
  pivot_wider(names_from = fin_year,
              values_from = rejected_samples_pct) %>% 
  rename("Health Board of Residence" = hbres_screen,
         "2020/21\u00b9" = "2020/21")

# Graph for printing
kpi2_2_graph <- cervical_column_graph(df = hist_rejected_table %>%
                                        filter(fin_year %in% fin_year_5) %>%
                                        select("Health Board of Residence" = hbres_screen,
                                               "Financial Year" = fin_year, 
                                               "Rejected (%)" = rejected_samples_pct),
                                      data_col = "Rejected (%)", 
                                      colours = create_palette(hist_rejected_table  |> 
                                                                 filter(fin_year %in% fin_year_5) |> 
                                                                 pull(fin_year)),
                                      threshold = 1, 
                                      pct = FALSE) # pct = FALSE so that the axis isn't showing up to 100%

ggsave(filename = paste0(fy_folder_path, "graphs/kpi2_2_", year_label, ".png"),
       plot = kpi2_2_graph,
       width = 9, height = 5)


## Write KPI 2 workbook ----
wrt_xlsx_kpi2(wb, scr_period_1yr, fmt_styles, kpi_notes$kpi2_notes, kpi2_list, 
              kpi2_1_table, kpi2_1_graph,
              kpi2_2_table, kpi2_2_graph,
              save_to_kpi2_xlsx)




# End of Script ----

