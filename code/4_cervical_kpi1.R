################################################################################
## R Cervical Screening Publication - KPI 1-3
## 4_cervical_kpi1.R
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
# Load in packages

if(is.na(utils::packageDate("pacman"))) install.packages("pacman")
pacman::p_load(phsmethods, readr, dplyr, tidyr, tidylog, lubridate, 
               forcats, ggplot2, openxlsx)

# Clean global environment
rm(list=ls())


# Source in other scripts containing functions and global variables
source("code/1_housekeeping.R")
source("src/func_cervical_kpi_calc_graph.R")
source("src/func_cervical_wrt_kpi_data.R")
source("src/func_cervical_wrt_xlsx_workbook.R")


# Clean global environment
rm(list=ls(pattern = "kpi[2-6]|atos|date"))

# 2: Read in extract data ----
### Import KPI extract
kpi1_1_extract <- arrow::read_parquet(paste0(fy_folder_path,
                                  "Checks/kpi1_1_extract", "_update", ".parquet")) %>%
  # remove any data without health board information
  filter(!is.na(`NHS Board of Residence`))


kpi1_2_extract <- arrow::read_parquet(paste0(fy_folder_path,
                                  "Checks/kpi1_2_extract", "_update", ".parquet")) %>%
  # remove any data without health board information
  filter(!is.na(`NHS Board of Residence`))

# 3: KPI 1 - Attendance and Uptake KPIs ----

## 3.1: KPI 1.1: Coverage of eligible individuals ---- 

### 3.1.1: KPI 1.1 Age ----
kpi1_1_25_64 <- kpi1_1_extract |> 
  calculate_kpi(num = screen, den = eligible, 
                age_group_25_64) |> 
  rename(age_group = age_group_25_64) |> 
  filter(!age_group %in% c("Under 25", "65+"))
  
kpi1_1_2549_5064 <- kpi1_1_extract |> 
  calculate_kpi(num = screen, den = eligible, 
                age_group2549_5064) |> 
  rename(age_group = age_group2549_5064) 

kpi1_1_5yr_grp <- kpi1_1_extract |> 
  calculate_kpi(num = screen, den = eligible, 
                age_group) |> 
  filter(!age_group %in% c("Under 25", "65+"))

kpi1_1 <- rbind(kpi1_1_25_64, kpi1_1_2549_5064, kpi1_1_5yr_grp) %>% 
  rename(screened = numerator,
         eligible = denominator,
         age_appr_coverage = percentage) %>% 
  mutate(hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                               levels = hbres_levels),
         .before = 1) %>% 
  arrange(hbres_screen)

### 3.1.2: KPI 1.1 SIMD ----
kpi1_1_simd_25_64 <- kpi1_1_extract |> 
  calculate_kpi(num = screen, den = eligible, 
                age_group_25_64, `Deprivation Category`) |> 
  rename(age_group = age_group_25_64) 

kpi1_1_simd_2549_5064 <- kpi1_1_extract |> 
  calculate_kpi(num = screen, den = eligible, 
                age_group2549_5064, `Deprivation Category`) |> 
  rename(age_group = age_group2549_5064) |> 
  filter(!age_group %in% c("Under 25", "65+"))

kpi1_1_simd_5yr_grp <- kpi1_1_extract |> 
  calculate_kpi(num = screen, den = eligible, 
                age_group, `Deprivation Category`) |> 
  filter(!age_group %in% c("Under 25", "65+"))


kpi1_1_simd <- rbind(kpi1_1_simd_25_64, kpi1_1_simd_2549_5064, kpi1_1_simd_5yr_grp) %>% 
  rename(simd = `Deprivation Category`,
         screened = numerator,
         eligible = denominator,
         age_appr_coverage = percentage) %>% 
  mutate(hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                               levels = hbres_levels),
         .before = 1) %>% 
  arrange(hbres_screen, simd)


## 3.2: KPI 1.2: Uptake of invited individuals ---- 

### 3.2.1: KPI 1.2 Age ----
kpi1_2_25_64 <- kpi1_2_extract |> 
  calculate_kpi(num = uptake_flag, den = invite_fy_flag, 
                age_group_25_64) |> 
  rename(age_group = age_group_25_64) |> 
  filter(!age_group %in% c("Under 25", "65+"))


kpi1_2_2549_5064 <- kpi1_2_extract |> 
  calculate_kpi(num = uptake_flag, den = invite_fy_flag, 
                age_group2549_5064) |> 
  rename(age_group = age_group2549_5064) |> 
  filter(!age_group %in% c("Under 25", "65+"))

kpi1_2_5yr_grp <- kpi1_2_extract |> 
  calculate_kpi(num = uptake_flag, den = invite_fy_flag, 
                age_group) |> 
  filter(!age_group %in% c("Under 25", "65+"))

kpi1_2_age <- rbind(kpi1_2_25_64, kpi1_2_2549_5064, kpi1_2_5yr_grp) %>% 
  rename("screened" = numerator,
         "invites" = denominator,
         "uptake" = percentage) %>% 
  mutate(#financial_year = fin_yr, 
    hbres_invite = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                          levels = hbres_levels),
    .before = 1) %>% 
  arrange(hbres_invite)

### 3.2.2: KPI 1.2 SIMD ----
kpi1_2_simd_25_64 <- kpi1_2_extract |> 
  calculate_kpi(num = uptake_flag, den = invite_fy_flag, 
                age_group_25_64, `Deprivation Category`) |> 
  rename(age_group = age_group_25_64) |> 
  filter(!age_group %in% c("Under 25", "65+"))

kpi1_2_simd_2549_5064 <- kpi1_2_extract |> 
  calculate_kpi(num = uptake_flag, den = invite_fy_flag, 
                age_group2549_5064, `Deprivation Category`) |> 
  rename(age_group = age_group2549_5064) |> 
  filter(!age_group %in% c("Under 25", "65+"))

kpi1_2_simd_5yr_grp <- kpi1_2_extract |> 
  calculate_kpi(num = uptake_flag, den = invite_fy_flag, 
                age_group, `Deprivation Category`) |> 
  filter(!age_group %in% c("Under 25", "65+"))

kpi1_2_simd <- rbind(kpi1_2_simd_25_64, kpi1_2_simd_2549_5064, kpi1_2_simd_5yr_grp) %>% 
  rename("simd" = `Deprivation Category`,
         "screened" = numerator,
         "invites" = denominator,
         "uptake" = percentage) %>%
  mutate(hbres_invite = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                               levels = hbres_levels),
         .before = 1) %>% 
  arrange(hbres_invite, simd)


# 4: Create Tables and Graphs ----

## 4.1: KPI 1 ----
### 4.1.1 - Coverage by HBs, 5 year trend data ----
cov_age_table <- kpi1_1 %>% 
  filter(age_group == "25-64") %>% 
  select(!c(hb2019, age_group, eligible, screened)) 

# Table for printing
kpi1_1_5yr <- cov_age_table %>% 
  filter(financial_year %in% fin_year_5) %>% 
  pivot_wider(names_from = financial_year,
              values_from = age_appr_coverage) %>% 
  rename("Health Board of Residence" = hbres_screen)

# Graph for printing
kpi1_1_graph <- cov_age_table %>%
  filter(financial_year %in% fin_year_5) %>%
  select("Health Board of Residence" = hbres_screen,
         "Financial Year" = financial_year, 
         "Age-appropriate Coverage (%)" = age_appr_coverage) |> 
  cervical_column_graph(data_col = "Age-appropriate Coverage (%)",
                        colours = create_palette(pull(cov_age_table, financial_year)),
    threshold = 80)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_1_", year_label, ".png"),
#        plot = kpi1_1_graph,
#        width = 9, height = 5)

# Scotland graph
kpi1_1_graph_scotland <- cov_age_table %>%
  filter(hbres_screen=="Scotland") %>%
  select("Health Board of Residence" = hbres_screen,
         "Financial Year" = financial_year, 
         "Age-appropriate Coverage (%)" = age_appr_coverage) |> 
  cervical_line_graph(data_col = "Age-appropriate Coverage (%)",
                      colours = create_palette(1),
                      threshold = 80
)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_1_", year_label, "_Scotland.png"),
#        plot = kpi1_1_graph_scotland,
#        width = 9, height = 5)

### 4.1.2: KPI 1 - Coverage by SIMD, 5 year trend data ----
simd_cov_table <- kpi1_1_simd %>% 
  filter(age_group == "25-64") %>% 
  select(!c(hb2019, age_group, eligible, screened)) %>% 
  mutate(simd = recode(simd, "1" = "1 - Most deprived", "5" = "5 - Least deprived"))

# Tables for printing
kpi1_1a_scotland <- simd_cov_table %>% 
  filter(financial_year %in% fin_year_5) %>% 
  filter(hbres_screen == "Scotland") 

kpi1_1a_5yr_table <- kpi1_1a_scotland %>% 
  pivot_wider(names_from = financial_year,
              values_from = age_appr_coverage)  %>% 
  select(!hbres_screen) %>% 
  rename("Deprivation (SIMD)" = simd)

kpi1_1a_1yr_table <- simd_cov_table %>% 
  filter(financial_year == fin_yr) %>% 
  select(!financial_year) %>% 
  pivot_wider(names_from = hbres_screen,
              values_from = age_appr_coverage) %>% 
  rename("Deprivation (SIMD)"= simd)

# Graph for printing
kpi1_1a_graph <- kpi1_1a_scotland %>%
  filter(financial_year %in% fin_year_5) %>%
  select("Deprivation (SIMD)" = simd, "Financial Year" = financial_year,
         "Age-appropriate Coverage (%)" = age_appr_coverage) |> 
  cervical_column_graph(data_col = "Age-appropriate Coverage (%)", 
                        colours = create_palette(pull(kpi1_1a_scotland, financial_year)),
                        threshold = 80)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_1a_", year_label, ".png"),
#        plot = kpi1_1a_graph,
#        width = 7.5, height = 4)

### 4.1.3: KPI 1 - Coverage by Age group, 5 year trend data ----
cov_age_table <- kpi1_1 %>% 
  filter(age_group %in% c("25-29", "30-34", "35-39", "40-44", 
                          "45-49", "50-54", "55-59", "60-64")) %>% 
  select(!c(hb2019, eligible, screened))

# Table for printing
kpi1_1b_scotland <- cov_age_table %>% 
  filter(financial_year %in% fin_year_5) %>% 
  filter(hbres_screen == "Scotland")

kpi1_1b_5yr_table <- kpi1_1b_scotland %>% 
  pivot_wider(names_from = financial_year,
              values_from = age_appr_coverage)  %>% 
  select(!hbres_screen) %>% 
  rename("Age Group" = age_group)

kpi1_1b_1yr_table <- cov_age_table %>% 
  filter(financial_year == fin_yr) %>% 
  select(!financial_year) %>% 
  pivot_wider(names_from = hbres_screen,
              values_from = age_appr_coverage) %>% 
  rename("Age Group" = age_group)


# Graph for printing
kpi1_1b_graph <- kpi1_1b_scotland %>%
  filter(financial_year %in% fin_year_5) %>%
  select("Age Group" = age_group, 
         "Financial Year" = financial_year, 
         "Age-appropriate Coverage (%)" = age_appr_coverage) |> 
  cervical_column_graph(data_col = "Age-appropriate Coverage (%)",
                        colours = create_palette(pull(kpi1_1b_scotland, financial_year)),
                        threshold = 80)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_1b_", year_label, ".png"),
#        plot = kpi1_1b_graph,
#        width = 7.5, height = 4)

### 4.1.4: KPI 1 - Coverage by SIMD and Age group, Scotland level, current financial year ----
kpi1_1c_scotland <- kpi1_1_simd %>% 
  filter(financial_year == fin_yr,
         hbres_screen == "Scotland") %>%  
  filter(age_group %in% c("25-29", "30-34", "35-39", "40-44", 
                          "45-49", "50-54", "55-59", "60-64")) %>% 
  select(!c(hb2019, eligible, screened)) %>% 
  mutate(simd = recode(simd, "1" = "1 - Most deprived", "5" = "5 - Least deprived"))

# Table for printing
kpi1_1c_table <- kpi1_1c_scotland %>% 
  pivot_wider(names_from = age_group,
              values_from = age_appr_coverage)  %>% 
  select(!c(hbres_screen, financial_year)) %>% 
  rename("Deprivation (SIMD)" = simd)

# Graph for printing
kpi1_1c_graph1 <- kpi1_1c_scotland %>%
  select("Age Group" = age_group,
         "Deprivation (SIMD)" = simd,
         "Age-appropriate Coverage (%)" = age_appr_coverage) |> 
  cervical_column_graph(data_col = "Age-appropriate Coverage (%)",
                        colours = create_palette(pull(kpi1_1c_scotland, simd)),
                        threshold = 80)

kpi1_1c_graph2 <- kpi1_1c_scotland %>%
  select("Deprivation (SIMD)" = simd,
         "Age Group" = age_group,
         "Age-appropriate Coverage (%)" = age_appr_coverage) |> 
  cervical_column_graph(data_col = "Age-appropriate Coverage (%)",
                        colours = create_palette(pull(kpi1_1c_scotland, age_group)),
                        threshold = 80)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_1c1_", year_label, ".png"),
#        plot = kpi1_1c_graph1,
#        width = 7.5, height = 4)
# ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_1c2_", year_label, ".png"),
#        plot = kpi1_1c_graph2,
#        width = 7.5, height = 4)


## 4.2: KPI 1.2 ----
### 4.2.1: - Uptake by HBs, 5 year trend data ----
upt_age_table <- kpi1_2_age %>% 
  filter(age_group == "25-64") %>% 
  select(!c(hb2019, age_group, invites, screened)) 

# Table for printing
kpi1_2_5yr <- upt_age_table %>% 
  filter(financial_year %in% fin_year_5) %>% 
  pivot_wider(names_from = financial_year,
              values_from = uptake) %>% 
  rename("Health Board of Residence" = hbres_invite)

# Graph for printing
kpi1_2_graph <- upt_age_table %>%
  filter(financial_year %in% fin_year_5) %>%
  select("Health Board of Residence" = hbres_invite,
         "Financial Year" = financial_year, 
         "Uptake (%)" = uptake) |> 
  cervical_column_graph(data_col = "Uptake (%)",
                        colours = create_palette(pull(upt_age_table, financial_year)),
                        threshold = 80)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_2_", year_label, ".png"),
#        plot = kpi1_2_graph,
#        width = 9, height = 5)

# Scotland Graph for Printing
kpi1_2_graph_scotland <- upt_age_table %>%
  filter(hbres_invite=="Scotland") %>%
  select("Health Board of Residence" = hbres_invite,
         "Financial Year" = financial_year, 
         "Uptake (%)" = uptake) |> 
  cervical_line_graph(data_col = "Uptake (%)",
                      colours = create_palette(1),
                      threshold = 80)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_2_", year_label, "_scotland.png"),
#        plot = kpi1_2_graph_scotland,
#        width = 9, height = 5)

### 4.2.2: KPI 1.2 - Uptake by SIMD, 5 year trend data ----
upt_simd_table <- kpi1_2_simd %>% 
  filter(age_group == "25-64") %>% 
  select(!c(hb2019, age_group, invites, screened)) %>% 
  mutate(simd = recode(simd, "1" = "1 - Most deprived", "5" = "5 - Least deprived"))

# Tables for printing
kpi1_2a_scotland <- upt_simd_table %>% 
  filter(financial_year %in% fin_year_5) %>% 
  filter(hbres_invite == "Scotland") 

kpi1_2a_5yr_table <- kpi1_2a_scotland %>% 
  pivot_wider(names_from = financial_year,
              values_from = uptake)  %>% 
  select(!hbres_invite) %>% 
  rename("Deprivation (SIMD)" = simd)

kpi1_2a_1yr_table <- upt_simd_table %>% 
  filter(financial_year == fin_yr) %>% 
  select(!financial_year) %>% 
  pivot_wider(names_from = hbres_invite,
              values_from = uptake) %>% 
  rename("Deprivation (SIMD)" = simd)

# Graph for printing
kpi1_2a_graph <- kpi1_2a_scotland %>%
  filter(financial_year %in% fin_year_5) %>% 
  select("Deprivation (SIMD)" = simd, "Financial Year" = financial_year,
         "Uptake (%)" = uptake) |> 
  cervical_column_graph(data_col = "Uptake (%)",
                        colours = create_palette(pull(kpi1_2a_scotland, financial_year)),
                        threshold = 80)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_2a_", year_label, ".png"),
#        plot = kpi1_2a_graph,
#        width = 7.5, height = 4)

### 4.2.3: KPI 1.2 - Coverage by Age group, 5 year trend data ----
upt_age_5yr_table <- kpi1_2_age %>%  
  filter(age_group %in% c("25-29", "30-34", "35-39", "40-44", 
                          "45-49", "50-54", "55-59", "60-64", "65+")) %>% 
  select(!c(hb2019, invites, screened))

# Table for printing
kpi1_2b_scotland <- upt_age_5yr_table %>% 
  filter(financial_year %in% fin_year_5) %>% 
  filter(hbres_invite == "Scotland")

kpi1_2b_5yr_table <- kpi1_2b_scotland %>% 
  select(!hbres_invite) %>% 
  pivot_wider(names_from = financial_year,
              values_from = uptake) %>% 
  rename("Age Group" = age_group)

kpi1_2b_1yr_table <- upt_age_5yr_table %>% 
  filter(financial_year == fin_yr) %>% 
  select(!financial_year) %>% 
  pivot_wider(names_from = hbres_invite,
              values_from = uptake) %>% 
  rename("Age Group" = age_group)


# Graph for printing
kpi1_2b_graph <- kpi1_2b_scotland %>%
  filter(financial_year %in% fin_year_5) %>% 
  select("Age Group" = age_group, 
         "Financial Year" = financial_year, 
         "Uptake (%)" = uptake) |> 
  cervical_column_graph(data_col = "Uptake (%)",
                        colours = create_palette(pull(kpi1_2b_scotland, financial_year)),
                        threshold = 80)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_2b_", year_label, ".png"),
#        plot = kpi1_2b_graph,
#        width = 7.5, height = 4)

### 4.2.4: KPI 1.2 - Uptake by SIMD and Age group, Scotland level, current financial year ----
kpi1_2c_scotland <- kpi1_2_simd %>% 
  filter(financial_year == fin_yr,
         hbres_invite == "Scotland") %>% 
  filter(age_group %in% c("25-29", "30-34", "35-39", "40-44", 
                          "45-49", "50-54", "55-59", "60-64", "65+")) %>% 
  select(!c(hb2019, invites, screened)) %>% 
  mutate(simd = recode(simd, "1" = "1 - Most deprived", "5" = "5 - Least deprived"))

# Table for printing
kpi1_2c_table <- kpi1_2c_scotland %>% 
  select(!c(hbres_invite, financial_year)) %>% 
  pivot_wider(names_from = age_group,
              values_from = uptake) %>% 
  rename("Deprivation (SIMD)" = simd)

# Graph for printing
kpi1_2c_graph1 <- kpi1_2c_scotland %>%
  select("Age Group" = age_group,
         "Deprivation (SIMD)" = simd,
         "Uptake (%)" = uptake) |> 
  cervical_column_graph(data_col = "Uptake (%)",
                        colours = create_palette(pull(kpi1_2c_scotland, simd)),
                        threshold = 80)

kpi1_2c_graph2 <- kpi1_2c_scotland %>%
  select("Deprivation (SIMD)" = simd,
         "Age Group" = age_group,
         "Uptake (%)" = uptake) |> 
  cervical_column_graph(data_col = "Uptake (%)", 
                        colours = create_palette(pull(kpi1_2c_scotland, age_group)),
                        threshold = 80)

# ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_2c1_", year_label, ".png"),
#        plot = kpi1_2c_graph1,
#        width = 7.5, height = 4)
# ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_2c2_", year_label, ".png"),
#        plot = kpi1_2c_graph2,
#        width = 7.5, height = 4)



# 5: Write KPI 1 workbook ----
wrt_xlsx_kpi1(scr_period_1yr, scr_period_5yr, fmt_styles, 
              kpi_notes$kpi1_notes, kpi1_list,
              kpi1_1_5yr, kpi1_1_graph_scotland, kpi1_1_graph,
              kpi1_1a_1yr_table, kpi1_1a_5yr_table, kpi1_1a_graph,
              kpi1_1b_1yr_table, kpi1_1b_5yr_table, kpi1_1b_graph,
              kpi1_1c_table, kpi1_1c_graph1, kpi1_1c_graph2,
             # kpi1_1d_table, kpi1_1d_graph,
              kpi1_2_5yr, kpi1_2_graph_scotland, kpi1_2_graph,
              kpi1_2a_1yr_table, kpi1_2a_5yr_table, kpi1_2a_graph,
              kpi1_2b_1yr_table, kpi1_2b_5yr_table, kpi1_2b_graph,
              kpi1_2c_table, kpi1_2c_graph1, kpi1_2c_graph2,
              #kpi1_2d_table, kpi1_2d_graph,
              save_to_kpi1_xlsx)


# End of Script ----
