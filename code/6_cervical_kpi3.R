##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## R Cervical Screening - KPIs for Publication
## 6_cervical_kpi3.R
## 
## Adult Screening Team
## Created: 30/08/24
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
rm(list=ls(pattern = "kpi[1-2]|fig[1-2]|simd"))

# Import extract ----
# If KPI analysis has previously been prepared for the financial year of interest
# Jump to section: Prepare data for publication
kpi3_extract <- read_rds(paste0(fy_folder_path,
                                "temp/kpi3_extract_", year_label, ".rds")) %>%
  # remove any data without health board information
  filter(!is.na(`NHS Board of Residence`))

kpi3_3_extract <- read_rds(paste0(fy_folder_path,
                                  "temp/kpi3_3_extract_", year_label, ".rds")) %>%
  # remove any data without health board information
  filter(!is.na(`NHS Board of Residence`))


# Import historic rds files ----
## KPI 3.1
hist_cyt_unsat <- readRDS(paste0(hist_folder_path, "/kpi3_1_hist_cytology_unsatis_", 
                                 substr(yr_start-1, 3, 4), substr(yr_end-1, 3, 4),
                                 ".rds"))

## KPI 3.2
hist_hpv_fail <- readRDS(paste0(hist_folder_path, "/kpi3_2_hist_hpv_failed_", 
                                substr(yr_start-1, 3, 4), substr(yr_end-1, 3, 4),
                                ".rds")) |> 
  filter(!fin_year %in% c("2016/17", "2017/18", "2018/19", "2019/20"))

## KPI 3.3
hist_repeat_test <- readRDS(paste0(hist_folder_path, "/kpi3_3_hist_repeat_test_", 
                                   substr(yr_start-1, 3, 4), substr(yr_end-1, 3, 4),
                                   ".rds"))


# KPI 3 - Screening Process KPIs ----
## KPI 3.1 - Percentage of unsatisfactory cytology tests
kpi3_1 <- cervical_pct(kpi3_extract, (reported_cyt_fy == 1), 
                       reported_cyt_unsat, reported_cyt_fy, hb2019) %>% 
  rename("cytology_samples" = Denominator,
         "unsatisfactory" = Numerator,
         "unsatisfactory_pct" = Percentage)

## KPI 3.2 - Percentage of tests processed which were reported as HPV fail
kpi3_2 <- cervical_pct(kpi3_extract, (reported_hpv_fy == 1),
                       reported_hpv_fail, reported_hpv_fy, hb2019) %>% 
  rename("virology_samples" = Denominator,
         "hpv_fail" = Numerator,
         "hpv_fail_pct" = Percentage)

## KPI 3.3 - Percentage of repeated tests due to cytology unsatisfactory or HPV fail
kpi3_3 <- cervical_pct(kpi3_3_extract, (first_reported_fy == 1),
                       repeat_scr, first_reported_fy,
                       hb2019) %>%
  rename("result_fail" = Denominator,
         "repeat_screen" = Numerator,
         "repeat_screen_pct" = Percentage)

## Save the new KPI 3 to historic records ----
hist_cyt_unsat_add <- rbind(hist_cyt_unsat, 
                            kpi3_1 %>%
                              mutate(fin_year = fin_yr, 
                                     hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                                                           levels = hbres_levels),
                                     .before = 1))

saveRDS(hist_cyt_unsat_add,
        file = paste0(fy_folder_path,
                      "historic/kpi3_1_hist_cytology_unsatis_", year_label, ".rds"))

hist_hpv_fail_add <- rbind(hist_hpv_fail, 
                           kpi3_2 %>%
                             mutate(fin_year = fin_yr, 
                                    hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                                                          levels = hbres_levels),
                                    .before = 1))

saveRDS(hist_hpv_fail_add, 
        file = paste0(fy_folder_path,
                      "historic/kpi3_2_hist_hpv_failed_", year_label, ".rds"))

hist_repeat_test_add <- rbind(hist_repeat_test, 
                              kpi3_3 %>%
                                mutate(fin_year = fin_yr, 
                                       hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                                                             levels = hbres_levels),
                                       .before = 1))

saveRDS(hist_repeat_test_add, 
        file = paste0(fy_folder_path,
                      "historic/kpi3_3_hist_repeat_test_", year_label, ".rds"))


# Prepare data for publication ----
# If above KPI data has previously been prepared, import rds files:
# hist_cyt_unsat_add <- read_rds(paste0(fy_folder_path,
#                                       "historic/kpi3_1_hist_cytology_unsatis_", year_label, ".rds"))
# hist_hpv_fail_add <- read_rds(paste0(fy_folder_path,
#                                      "historic/kpi3_2_hist_hpv_failed_", year_label, ".rds"))
# hist_repeat_test_add <- read_rds(paste0(fy_folder_path,
#                                         "historic/kpi3_3_hist_repeat_test_", year_label, ".rds"))

## KPI 3.1 - Percentage of unsatisfactory cytology tests ----
hist_cyt_unsat_table <- hist_cyt_unsat_add %>% 
  select(!c(hb2019, cytology_samples, unsatisfactory))

# Table for printing
kpi3_1_table <- hist_cyt_unsat_table %>% 
  #filter(fin_year %in% fin_year_5) %>% # add back in after 2022/23 and 2023/24 combined publication 
  pivot_wider(names_from = fin_year,
              values_from = unsatisfactory_pct) %>% 
  rename("Health Board of Residence" = hbres_screen,
         "2020/21\u00b9" = "2020/21")

# Graph for printing
kpi3_1_graph <- cervical_column_graph(df = hist_cyt_unsat_table %>%
                                        filter(fin_year %in% fin_year_5) %>%
                                        select("Health Board of Residence" = hbres_screen,
                                               "Financial Year" = fin_year, 
                                               "Cytology Unsatisfactory (%)" = unsatisfactory_pct),
                                      data_col = "Cytology Unsatisfactory (%)",
                                      colours = create_palette(hist_cyt_unsat_table |> 
                                                                 filter(fin_year %in% fin_year_5) |> 
                                                                 pull(fin_year)), 
                                      threshold = 4, 
                                      pct = FALSE)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi3_1_", year_label, ".png"),
       plot = kpi3_1_graph,
       width = 9, height = 5)

## KPI 3.2 - Percentage of tests processed which were reported as HPV fail ----
hist_hpv_fail_table <- hist_hpv_fail_add %>% 
  select(!c(hb2019, virology_samples, hpv_fail))

# Table for printing
kpi3_2_table <- hist_hpv_fail_table %>% 
  #filter(fin_year %in% fin_year_5) %>% # add back in after 2022/23 and 2023/24 combined publication 
  pivot_wider(names_from = fin_year,
              values_from = hpv_fail_pct) %>% 
  rename("Health Board of Residence" = hbres_screen,
         "2020/21\u00b9" = "2020/21")

# Graph for printing
kpi3_2_graph <- cervical_column_graph(df = hist_hpv_fail_table %>%
                                        filter(fin_year %in% fin_year_5) %>%
                                        select("Health Board of Residence" = hbres_screen,
                                               "Financial Year" = fin_year, 
                                               "HPV Fail (%)" = hpv_fail_pct),
                                      data_col = "HPV Fail (%)", 
                                      colours = create_palette(hist_hpv_fail_table |> 
                                                                 filter(fin_year %in% fin_year_5) |> 
                                                                 pull(fin_year)), 
                                      threshold = 1, 
                                      pct = FALSE)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi3_2_", year_label, ".png"),
       plot = kpi3_2_graph,
       width = 9, height = 5)

## KPI 3.3 - Percentage of repeated tests due to cytology unsatisfactory or HPV fail ----
hist_repeat_test_table <- hist_repeat_test_add %>% 
  select(!c(hb2019, result_fail, repeat_screen))

# Table for printing
kpi3_3_table <- hist_repeat_test_table %>% 
  #filter(fin_year %in% fin_year_5) %>% # add back in after 2022/23 and 2023/24 combined publication 
  pivot_wider(names_from = fin_year,
              values_from = repeat_screen_pct) %>% 
  rename("Health Board of Residence" = hbres_screen,
         "2020/21\u00b9" = "2020/21")

# Graph for printing
kpi3_3_graph <- cervical_column_graph(df = hist_repeat_test_table %>%
                                        filter(fin_year %in% fin_year_5) %>%
                                        select("Health Board of Residence" = hbres_screen,
                                               "Financial Year" = fin_year, 
                                               "Repeat Screen (%)" = repeat_screen_pct),
                                      data_col = "Repeat Screen (%)",
                                      colours = create_palette(hist_repeat_test_table |> 
                                                                 filter(fin_year %in% fin_year_5) |> 
                                                                 pull(fin_year)), 
                                      threshold = 100)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi3_3_", year_label, ".png"),
       plot = kpi3_3_graph,
       width = 9, height = 5)

## Supplementary data for number of total samples processed ----
kpi3_total_samples <- hist_cyt_unsat_add |> 
  left_join(hist_hpv_fail_add, by = c("fin_year", "hbres_screen", "hb2019")) |> 
  filter(hbres_screen == "Scotland") |> 
  select("Financial Year" = fin_year, 
         "Cytology Samples" = cytology_samples,
         "Virology Samples" = virology_samples) |> 
  pivot_longer(cols = c(2:3),
               names_to = "Sample Type",
               values_to = "number") |> 
  pivot_wider(names_from = 1,
              values_from = number) |> 
  rename("2020/21\u00b9" = "2020/21")

## Write KPI 3 workbook ----
wrt_xlsx_kpi3(wb, scr_period_1yr, fmt_styles, kpi_notes$kpi3_notes, kpi3_list, 
              kpi3_total_samples, 
              kpi3_1_table, kpi3_1_graph,
              kpi3_2_table, kpi3_2_graph,
              kpi3_3_table, kpi3_3_graph,
              save_to_kpi3_xlsx)



# End of Script ----
