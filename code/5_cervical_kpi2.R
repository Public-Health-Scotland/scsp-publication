################################################################################
## R Cervical Screening Publication - KPI 1-3
## 5_cervical_kpi2.R
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

# Packages ----
if(is.na(utils::packageDate("pacman"))) install.packages("pacman")
pacman::p_load(phsmethods, readr, dplyr, tidyr, tidylog, janitor, 
               lubridate, forcats, ggplot2, openxlsx)

## Clean global environment
rm(list=ls())


# Source in other scripts containing functions and global variables
source("code/1_housekeeping.R")
source("src/func_cervical_kpi_calc_graph.R")
source("src/func_cervical_wrt_kpi_data.R")
source("src/func_cervical_wrt_xlsx_workbook.R")

# Clean global environment
rm(list=ls(pattern = "kpi1|kpi[3-6]|fig1|fig[3-6]|atos|simd"))


# 2: Import extract data ----
kpi2_extract <- arrow::read_parquet(paste0(fy_folder_path,
                                             "temp/kpi2_extract", "_update", ".parquet")) %>%
  # remove any data without health board information
  filter(!is.na(`NHS Board of Residence`))

# 3: KPI 2 - Laboratory Service KPIs ----

## 3.1: KPI 2.1 ----
# Annual percentage of samples reported within 10 working days from the date 
# of the sample being taken

kpi2_1 <- kpi2_extract |> 
  calculate_kpi(num = reported_14, den = reported_fy) %>% 
  rename("total_reports" = denominator,
         "report_2week" = numerator,
         "report_2week_pct" = percentage) %>%
  mutate(hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                          levels = hbres_levels),
    .before = 1) %>% 
  arrange(hbres_screen)


## 3.2: KPI 2.2 ----
# Percentage of tests rejected by the laboratory prior to processing KPI 2 
# document requests data to be broken down by rejection reason. This is not done 
# in the primary data table, but presented in a separate data table due to there 
# being so few rejected samples

kpi2_2 <- kpi2_extract |> 
  calculate_kpi(num = rejected, den = arrived_fy) %>% 
  rename("total_samples" = denominator,
         "rejected_samples" = numerator,
         "rejected_samples_pct" = percentage)%>%
  mutate(hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                          levels = hbres_levels),
    .before = 1) %>% 
  arrange(hbres_screen)

# 4: Create Tables and Graphs ----
## 4.1: KPI 2.1 - Number of samples reported within 2-weeks ----
report_table <- kpi2_1 %>% 
  select(!c(hb2019, total_reports, report_2week))

# Table for printing
kpi2_1_table <- report_table %>% 
  filter(financial_year %in% fin_year_5) %>% # add back in after 2022/23 and 2023/24 combined publication 
  pivot_wider(names_from = financial_year,
              values_from = report_2week_pct) %>% 
  rename("Health Board of Residence" = hbres_screen)|>
  rename(any_of(c("2020/21\u00b9" = "2020/21")))

# Graph for printing
kpi2_1_graph <- report_table %>%
  filter(financial_year %in% fin_year_5) %>%
  select("Health Board of Residence" = hbres_screen,
         "Financial Year" = financial_year, 
         "Reported within 2-weeks (%)" = report_2week_pct) |> 
  cervical_column_graph(data_col = "Reported within 2-weeks (%)",
                        colours = create_palette(pull(report_table, financial_year)),
                        threshold = 80)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi2_1_", year_label, ".png"),
#        plot = kpi2_1_graph,
#        width = 9, height = 5)

## 4.2: KPI 2.2 - Number of samples rejected by lab ----
rejected_table <- kpi2_2 %>% 
  select(!c(hb2019, total_samples, rejected_samples))

# Table for printing
kpi2_2_table <- rejected_table %>% 
  filter(financial_year %in% fin_year_5) %>% # add back in after 2022/23 and 2023/24 combined publication 
  pivot_wider(names_from = financial_year,
              values_from = rejected_samples_pct) %>% 
  rename("Health Board of Residence" = hbres_screen,
         any_of(c("2020/21\u00b9" = "2020/21")))

# Graph for printing
kpi2_2_graph <- rejected_table %>%
  filter(financial_year %in% fin_year_5) %>%
  select("Health Board of Residence" = hbres_screen,
         "Financial Year" = financial_year, 
         "Rejected (%)" = rejected_samples_pct) |> 
  cervical_column_graph(data_col = "Rejected (%)",
                        colours = create_palette(pull(rejected_table, financial_year)),
                        threshold = 1,
                        pct = FALSE) # pct = FALSE so that the axis isn't showing up to 100%

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi2_2_", year_label, ".png"),
#        plot = kpi2_2_graph,
#        width = 9, height = 5)


# 5: Write KPI 2 workbook ----
wrt_xlsx_kpi2(wb, scr_period_1yr, fmt_styles, kpi_notes$kpi2_notes, kpi2_list, 
              kpi2_1_table, kpi2_1_graph,
              kpi2_2_table, kpi2_2_graph,
              save_to_kpi2_xlsx)

#### End of Script ####
