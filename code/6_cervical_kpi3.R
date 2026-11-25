################################################################################
## R Cervical Screening Publication - KPI 1-3
## 6_cervical_kpi3.R
## 
## Adult Screening Team
## Created: 23/08/2024
## Last update: 25/11/2026
## Analysts: Ben Duncan & John Quinn
## 
## Written/run on RStudio Posit Workbench
## Platform: x86_64-pc-linux-gnu
## R version 4.4.2 (Pile of Leaves) - 2024-10-31
################################################################################

# Load in packages
if(is.na(utils::packageDate("pacman"))) install.packages("pacman")
pacman::p_load(phsmethods, readr, dplyr, tidyr, tidylog, janitor, 
               lubridate, forcats, ggplot2, openxlsx)

# Clean global environment
rm(list=ls())

# Source in other scripts containing functions and global variables
source("code/1_housekeeping.R")
source("src/func_cervical_kpi_calc_graph.R")
source("src/func_cervical_wrt_kpi_data.R")
source("src/func_cervical_wrt_xlsx_workbook.R")

# Clean global environment
rm(list=ls(pattern = "kpi[1-2]|kpi[4-6]|fig[1-2]|fig[4-6]|atos|simd"))

# 2: Import Extract Dara ----
kpi3_extract <- arrow::read_parquet(paste0(fy_folder_path,
                                             "temp/kpi3_extract", "_update", ".parquet")) %>%
  # remove any data without health board information
  filter(!is.na(`NHS Board of Residence`))

kpi3_3_extract <- arrow::read_parquet(paste0(fy_folder_path,
                                             "temp/kpi3_3_extract", "_update", ".parquet")) %>%
  # remove any data without health board information
  filter(!is.na(`NHS Board of Residence`))


# 3: KPI 3 - Screening Process KPIs ----

## 3.1: KPI 3.1 - Percentage of unsatisfactory cytology tests ----
kpi3_1 <- kpi3_extract |> 
  calculate_kpi(num = reported_cyt_unsat, den = reported_cyt_fy) %>% 
  rename("cytology_samples" = denominator,
         "unsatisfactory" = numerator,
         "unsatisfactory_pct" = percentage) |> 
  mutate(hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                          levels = hbres_levels),
    .before = 1)



## 3.2: KPI 3.2 - Percentage of tests processed which were reported as HPV fail ----
kpi3_2 <- kpi3_extract |> 
  calculate_kpi(num = reported_hpv_fail, den = reported_hpv_fy) %>% 
  rename("virology_samples" = denominator,
         "hpv_fail" = numerator,
         "hpv_fail_pct" = percentage)%>%
  mutate(hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                          levels = hbres_levels),
    .before = 1) |> 
  filter(!financial_year %in% c("2016/17", "2017/18", "2018/19", "2019/20"))


## 3.3: KPI 3.3 - Percentage of repeated tests due to cytology unsatisfactory or HPV fail ----

## Numerator: Number of individuals who are screening following cytology unsatisfactory or HPV fail
## Denominator: Total number of individuals with an unsatisfactory result annually
kpi3_3 <- kpi3_3_extract |> 
  calculate_kpi(num = repeat_scr, den = first_reported_fy) %>% 
  rename("result_fail" = denominator,
         "repeat_screen" = numerator,
         "repeat_screen_pct" = percentage) |> 
  mutate(hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                                                             levels = hbres_levels),
         .before = 1)


# 4: Create Tables and Graphs ----
## 4.1: KPI 3.1 - Percentage of unsatisfactory cytology tests ----
cyt_unsat_table <- kpi3_1 %>% 
  select(!c(hb2019, cytology_samples, unsatisfactory))

# Table for printing
kpi3_1_table <- cyt_unsat_table %>% 
  filter(financial_year %in% fin_year_5) %>% # add back in after 2022/23 and 2023/24 combined publication 
  pivot_wider(names_from = financial_year,
              values_from = unsatisfactory_pct) %>% 
  rename("Health Board of Residence" = hbres_screen,
         any_of(c("2020/21\u00b9" = "2020/21")))

# Graph for printing
kpi3_1_graph <- cyt_unsat_table %>%
  filter(financial_year %in% fin_year_5) %>%
  select("Health Board of Residence" = hbres_screen,
         "Financial Year" = financial_year, 
         "Cytology Unsatisfactory (%)" = unsatisfactory_pct) |> 
  cervical_column_graph(data_col = "Cytology Unsatisfactory (%)",
                        colours = create_palette(pull(cyt_unsat_table ,financial_year)), 
                                      threshold = 4, 
                                      pct = FALSE)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi3_1_", year_label, ".png"),
#        plot = kpi3_1_graph,
#        width = 9, height = 5)

## 4.2: KPI 3.2 - Percentage of tests processed which were reported as HPV fail ----
hpv_fail_table <- kpi3_2 %>% 
  select(!c(hb2019, virology_samples, hpv_fail))

# Table for printing
kpi3_2_table <- hpv_fail_table %>% 
  filter(financial_year %in% fin_year_5) %>% # add back in after 2022/23 and 2023/24 combined publication 
  pivot_wider(names_from = financial_year,
              values_from = hpv_fail_pct) %>% 
  rename("Health Board of Residence" = hbres_screen,
         any_of(c("2020/21\u00b9" = "2020/21")))

# Graph for printing
kpi3_2_graph <- hpv_fail_table %>%
  filter(financial_year %in% fin_year_5) %>%
   select("Health Board of Residence" = hbres_screen,
          "Financial Year" = financial_year, 
          "HPV Fail (%)" = hpv_fail_pct) |> 
  cervical_column_graph(data_col = "HPV Fail (%)",
                        colours = create_palette(pull(hpv_fail_table, financial_year)), 
                                      threshold = 1, 
                                      pct = FALSE)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi3_2_", year_label, ".png"),
#        plot = kpi3_2_graph,
#        width = 9, height = 5)

## 4.3: KPI 3.3 - Percentage of repeated tests due to cytology unsatisfactory or HPV fail ----
repeat_test_table <- kpi3_3 %>% 
  select(!c(hb2019, result_fail, repeat_screen))

# Table for printing
kpi3_3_table <- repeat_test_table %>% 
  filter(financial_year %in% fin_year_5) %>% # add back in after 2022/23 and 2023/24 combined publication 
  pivot_wider(names_from = financial_year,
              values_from = repeat_screen_pct) %>% 
  rename("Health Board of Residence" = hbres_screen,
         any_of(c("2020/21\u00b9" = "2020/21")))

# Graph for printing
kpi3_3_graph <- repeat_test_table %>%
  filter(financial_year %in% fin_year_5) %>%
  select("Health Board of Residence" = hbres_screen,
         "Financial Year" = financial_year, 
         "Repeat Screen (%)" = repeat_screen_pct) |> 
  cervical_column_graph(data_col = "Repeat Screen (%)",
                        colours = create_palette(pull(repeat_test_table, financial_year)), 
                                      threshold = 100)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi3_3_", year_label, ".png"),
#        plot = kpi3_3_graph,
#        width = 9, height = 5)

## 4.4: Supplementary data for number of total samples processed ----
kpi3_total_samples <- kpi3_1 |> 
  left_join(kpi3_2, by = c("financial_year", "hbres_screen", "hb2019")) |> 
  filter(hbres_screen == "Scotland") |> 
  select("Financial Year" = financial_year, 
         "Cytology Samples" = cytology_samples,
         "Virology Samples" = virology_samples) |> 
  pivot_longer(cols = c(2:3),
               names_to = "Sample Type",
               values_to = "number") |> 
  pivot_wider(names_from = 1,
              values_from = number) |> 
  rename(any_of(c("2020/21\u00b9" = "2020/21")))

# 5: Write KPI 3 workbook ----
wrt_xlsx_kpi3(wb, scr_period_1yr, fmt_styles, kpi_notes$kpi3_notes, kpi3_list, 
              kpi3_total_samples, 
              kpi3_1_table, kpi3_1_graph,
              kpi3_2_table, kpi3_2_graph,
              kpi3_3_table, kpi3_3_graph,
              save_to_kpi3_xlsx)

# End of Script ----
