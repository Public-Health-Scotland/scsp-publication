##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## R Cervical Screening - KPIs for Publication
## 4_cervical_kpi1.R
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

## Clean global environment
rm(list=ls(pattern = "kpi[2-3]|date"))


# Import KPI extract ----
# If KPI analysis has previously been prepared for the financial year of interest
# Jump to section: Prepare data for publication
kpi1_1_extract <- read_rds(paste0(fy_folder_path,
                                  "temp/kpi1_1_extract_", year_label, ".rds")) %>%
  # remove any data without health board information
  filter(!is.na(`NHS Board of Residence`))

kpi1_2_extract <- read_rds(paste0(fy_folder_path,
                                  "temp/kpi1_2_extract_", year_label, ".rds")) %>%
  # remove any data without health board information
  filter(!is.na(`NHS Board of Residence`))


# Import historic rds files ----
## KPI 1.1
hist_cov <- readRDS(paste0(hist_folder_path, "/kpi1_1_hist_coverage_", 
                           substr(yr_start-1, 3, 4), substr(yr_end-1, 3, 4),
                           ".rds"))

hist_simd_cov <- readRDS(paste0(hist_folder_path, "/kpi1_1_hist_coverage_simd_",
                                substr(yr_start-1, 3, 4), substr(yr_end-1, 3, 4),
                                ".rds"))

## KPI 1.2
hist_upt <- readRDS(paste0(hist_folder_path, "/kpi1_2_hist_uptake_", 
                           substr(yr_start-1, 3, 4), substr(yr_end-1, 3, 4),
                           ".rds"))

hist_simd_upt <- readRDS(paste0(hist_folder_path, "/kpi1_2_hist_uptake_simd_",
                                substr(yr_start-1, 3, 4), substr(yr_end-1, 3, 4),
                                ".rds"))


# KPI 1 - Coverage and Uptake KPIs ----
## KPI 1.1: Coverage of eligible individuals who have attended screening ----
## within a 3.5 and/or 5.5-year reporting period broken down by age and SIMD
kpi1_1_25_64 <- cervical_pct(kpi1_1_extract, (eligible == 1),
                       scr5_5, eligible,
                       hb2019, age_group_25_64,
                       calc_coverage = TRUE) %>% 
  rename("age_group" = age_group_25_64) %>% 
  filter(age_group != "Under 25")

kpi1_1_2549_5064 <- cervical_pct(kpi1_1_extract, (eligible == 1),
                                 scr5_5, eligible,
                                 hb2019, age_group2549_5064,
                                 calc_coverage = TRUE) %>% 
  rename("age_group" = age_group2549_5064)

kpi1_1_5yr_grp <- cervical_pct(kpi1_1_extract, (eligible == 1),
                               scr5_5, eligible,
                               hb2019, age_group,
                               calc_coverage = TRUE) %>% 
  filter(age_group != "Under 25")

kpi1_1_simd_25_64 <- cervical_pct(kpi1_1_extract, (eligible == 1),
                       scr5_5, eligible,
                       hb2019, age_group_25_64, `Deprivation Category`,
                       calc_coverage = TRUE) %>% 
  rename("age_group" = age_group_25_64)

kpi1_1_simd_2549_5064 <- cervical_pct(kpi1_1_extract, (eligible == 1),
                        scr5_5, eligible,
                        hb2019, age_group2549_5064, `Deprivation Category`,
                        calc_coverage = TRUE) %>% 
  rename("age_group" = age_group2549_5064) %>% 
  filter(age_group != "Under 25")

kpi1_1_simd_5yr_grp <- cervical_pct(kpi1_1_extract, (eligible == 1),
                               scr5_5, eligible,
                               hb2019, age_group, `Deprivation Category`,
                               calc_coverage = TRUE) %>% 
  filter(age_group != "Under 25")


## KPI 1.2: Annual uptake among individuals who are sent an offer of screening ----
## in the latest financial year broken down by age and SIMD
kpi1_2_25_64 <- cervical_pct(kpi1_2_extract, (invite_fy_flag == 1),
                             uptake_flag, invite_fy_flag,
                             hb2019, age_group_25_64) %>% 
  rename("age_group" = age_group_25_64) %>% 
  filter(age_group != "65+")

kpi1_2_2549_5064 <- cervical_pct(kpi1_2_extract, (invite_fy_flag == 1),
                                 uptake_flag, invite_fy_flag,
                                 hb2019, age_group2549_5064) %>% 
  rename("age_group" = age_group2549_5064) %>% 
  filter(age_group != "65+")

kpi1_2_5yr_grp <- cervical_pct(kpi1_2_extract, (invite_fy_flag == 1),
                               uptake_flag, invite_fy_flag,
                               hb2019, age_group)

kpi1_2_simd_25_64 <- cervical_pct(kpi1_2_extract, (invite_fy_flag == 1),
                                  uptake_flag, invite_fy_flag,
                                  hb2019, age_group_25_64, `Deprivation Category`) %>% 
  rename("age_group" = age_group_25_64) %>% 
  filter(age_group != "65+")

kpi1_2_simd_2549_5064 <- cervical_pct(kpi1_2_extract, (invite_fy_flag == 1),
                                      uptake_flag, invite_fy_flag,
                                      hb2019, age_group2549_5064, `Deprivation Category`) %>% 
  rename("age_group" = age_group2549_5064) %>% 
  filter(age_group != "65+")

kpi1_2_simd_5yr_grp <- cervical_pct(kpi1_2_extract, (invite_fy_flag == 1),
                               uptake_flag, invite_fy_flag,
                               hb2019, age_group, `Deprivation Category`)


# Save the new KPI 1 to historic records ----
## KPI 1.1 - Combine data frames to get all age group for historic data ----
### KPI 1.1
kpi1_1_age <- rbind(kpi1_1_25_64, kpi1_1_2549_5064, kpi1_1_5yr_grp) %>% 
  rename("screened" = Numerator,
         "eligible" = Denominator,
         "age_appr_coverage" = Percentage) %>% 
  mutate(fin_year = fin_yr, 
         hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                               levels = hbres_levels),
         .before = 1) %>% 
  arrange(hbres_screen)

hist_cov_add <- rbind(hist_cov, 
                      kpi1_1_age)

saveRDS(hist_cov_add, 
        file = paste0(fy_folder_path,
                      "historic/kpi1_1_hist_coverage_", year_label, ".rds"))


### KPI 1.1a - SIMD
kpi1_1_simd <- rbind(kpi1_1_simd_25_64, kpi1_1_simd_2549_5064, kpi1_1_simd_5yr_grp) %>% 
  rename("simd" = `Deprivation Category`,
         "screened" = Numerator,
         "eligible" = Denominator,
         "age_appr_coverage" = Percentage) %>%
  mutate(fin_year = fin_yr, 
         hbres_screen = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                               levels = hbres_levels),
         .before = 1) %>% 
  arrange(hbres_screen, simd)

hist_simd_cov_add <- rbind(hist_simd_cov, 
                           kpi1_1_simd)

saveRDS(hist_simd_cov_add, 
        file = paste0(fy_folder_path,
                      "historic/kpi1_1_hist_coverage_simd_", year_label, ".rds"))


## KPI 1.2 - Combine data frames to get all age group for historic data ----
### KPI 1.2
kpi1_2_age <- rbind(kpi1_2_25_64, kpi1_2_2549_5064, kpi1_2_5yr_grp) %>% 
  rename("screened" = Numerator,
         "invites" = Denominator,
         "uptake" = Percentage) %>% 
  mutate(fin_year = fin_yr, 
         hbres_invite = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                               levels = hbres_levels),
         .before = 1) %>% 
  arrange(hbres_invite)
# Until we figure out what we need, all HBs and age groups are included in the historical coverage file

hist_upt_add <- rbind(hist_upt, 
                      kpi1_2_age)

saveRDS(hist_upt_add, 
        file = paste0(fy_folder_path,
                      "historic/kpi1_2_hist_uptake_", year_label, ".rds"))


### KPI 1.2a - SIMD
kpi1_2_simd <- rbind(kpi1_2_simd_25_64, kpi1_2_simd_2549_5064, kpi1_2_simd_5yr_grp) %>% 
  rename("simd" = `Deprivation Category`,
         "screened" = Numerator,
         "invites" = Denominator,
         "uptake" = Percentage) %>%
  mutate(fin_year = fin_yr, 
         hbres_invite = factor(stringr::str_replace(match_area(as.character(hb2019)), " and ", " & "),
                               levels = hbres_levels),
         .before = 1) %>% 
  arrange(hbres_invite, simd)

hist_simd_upt_add <- rbind(hist_simd_upt, 
                           kpi1_2_simd)

saveRDS(hist_simd_upt_add, 
        file = paste0(fy_folder_path,
                      "historic/kpi1_2_hist_uptake_simd_", year_label, ".rds"))


# Prepare data for publication ----
# If above KPI data has previously been prepared, import rds files:
# hist_cov_add <- readRDS(paste0(fy_folder_path, "historic/kpi1_1_hist_coverage_",
#                                year_label, ".rds"))
# 
# hist_simd_cov_add <- readRDS(paste0(fy_folder_path, "historic/kpi1_1_hist_coverage_simd_",
#                                     year_label, ".rds"))
# 
# ## KPI 1.2
# hist_upt_add <- readRDS(paste0(fy_folder_path, "historic/kpi1_2_hist_uptake_",
#                                year_label, ".rds"))
# 
# hist_simd_upt_add <- readRDS(paste0(fy_folder_path, "historic/kpi1_2_hist_uptake_simd_",
#                                     year_label, ".rds"))

## KPI 1 - Coverage by HBs, 5 year trend data ----
hist_cov_table <- hist_cov_add %>% 
  filter(age_group == "25-64") %>% 
  select(!c(hb2019, age_group, eligible, screened)) 

# Table for printing
kpi1_1_5yr <- hist_cov_table %>% 
  #filter(fin_year %in% fin_year_5) %>% 
  pivot_wider(names_from = fin_year,
              values_from = age_appr_coverage) %>% 
  rename("Health Board of Residence" = hbres_screen)

# Graph for printing
kpi1_1_graph <- cervical_column_graph(df = hist_cov_table %>%
                                        filter(fin_year %in% fin_year_5) %>%
                                        select("Health Board of Residence" = hbres_screen,
                                               "Financial Year" = fin_year, 
                                               "Age-appropriate Coverage (%)" = age_appr_coverage),
                                      data_col = "Age-appropriate Coverage (%)", 
                                      colours = create_palette(hist_cov_table |> 
                                                                 filter(fin_year %in% fin_year_5) |> 
                                                                 pull(fin_year)),
                                      threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_1_", year_label, ".png"),
       plot = kpi1_1_graph,
       width = 9, height = 5)

# Scotland graph
kpi1_1_graph_scotland <- cervical_line_graph(df = hist_cov_table %>%
                                               filter(hbres_screen=="Scotland") %>%
                                               select("Health Board of Residence" = hbres_screen,
                                                      "Financial Year" = fin_year, 
                                                      "Age-appropriate Coverage (%)" = age_appr_coverage),
                                             data_col <- "Age-appropriate Coverage (%)",
                                             colours <- create_palette(1),
                                             threshold = 80
                                             )

ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_1_", year_label, "_Scotland.png"),
       plot = kpi1_1_graph_scotland,
       width = 9, height = 5)


## KPI 1 - Coverage by SIMD, 5 year trend data ----
hist_simd_cov_table <- hist_simd_cov_add %>% 
  filter(age_group == "25-64") %>% 
  select(!c(hb2019, age_group, eligible, screened)) %>% 
  mutate(simd = recode(simd, "1" = "1 - Most deprived", "5" = "5 - Least deprived"))

# Tables for printing
kpi1_1a_scotland <- hist_simd_cov_table %>% 
  #filter(fin_year %in% fin_year_5) %>% 
  filter(hbres_screen == "Scotland") 

kpi1_1a_5yr_table <- kpi1_1a_scotland %>% 
  pivot_wider(names_from = fin_year,
              values_from = age_appr_coverage)  %>% 
  select(!hbres_screen) %>% 
  rename("Deprivation (SIMD)" = simd)

kpi1_1a_1yr_table <- hist_simd_cov_table %>% 
  filter(fin_year == fin_yr) %>% 
  select(!fin_year) %>% 
  pivot_wider(names_from = hbres_screen,
              values_from = age_appr_coverage) %>% 
  rename("Deprivation (SIMD)"= simd)

# Graph for printing
kpi1_1a_graph <- cervical_column_graph(df = kpi1_1a_scotland %>%
                                         filter(fin_year %in% fin_year_5) %>%
                                         select("Deprivation (SIMD)" = simd, "Financial Year" = fin_year,
                                                "Age-appropriate Coverage (%)" = age_appr_coverage),
                                       data_col = "Age-appropriate Coverage (%)", 
                                       colours = create_palette(kpi1_1a_scotland %>%
                                                                  filter(fin_year %in% fin_year_5) |> 
                                                                  pull(fin_year)),
                                       threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_1a_", year_label, ".png"),
       plot = kpi1_1a_graph,
       width = 7.5, height = 4)

## KPI 1 - Coverage by Age group, 5 year trend data ----
hist_cov_age_table <- hist_cov_add %>% 
  filter(age_group %in% c("25-29", "30-34", "35-39", "40-44", 
                          "45-49", "50-54", "55-59", "60-64")) %>% 
  select(!c(hb2019, eligible, screened))

# Table for printing
kpi1_1b_scotland <- hist_cov_age_table %>% 
  #filter(fin_year %in% fin_year_5) %>% 
  filter(hbres_screen == "Scotland")

kpi1_1b_5yr_table <- kpi1_1b_scotland %>% 
  pivot_wider(names_from = fin_year,
              values_from = age_appr_coverage)  %>% 
  select(!hbres_screen) %>% 
  rename("Age Group" = age_group)

kpi1_1b_1yr_table <- hist_cov_age_table %>% 
  filter(fin_year == fin_yr) %>% 
  select(!fin_year) %>% 
  pivot_wider(names_from = hbres_screen,
              values_from = age_appr_coverage) %>% 
  rename("Age Group" = age_group)

# Graph for printing
kpi1_1b_graph <- cervical_column_graph(df = kpi1_1b_scotland %>%
                                         filter(fin_year %in% fin_year_5) %>%
                                         select("Age Group" = age_group, 
                                                "Financial Year" = fin_year, 
                                                "Age-appropriate Coverage (%)" = age_appr_coverage),
                                       data_col = "Age-appropriate Coverage (%)", 
                                       colours = create_palette(kpi1_1b_scotland %>%
                                                                  filter(fin_year %in% fin_year_5) |> 
                                                                  pull(fin_year)),
                                       threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_1b_", year_label, ".png"),
       plot = kpi1_1b_graph,
       width = 7.5, height = 4)


## KPI 1 - Coverage by SIMD and Age group, Scotland level, current financial year ----
kpi1_1c_scotland <- hist_simd_cov_add %>% 
  filter(fin_year == fin_yr,
         hbres_screen == "Scotland") %>%  
  filter(age_group %in% c("25-29", "30-34", "35-39", "40-44", 
                          "45-49", "50-54", "55-59", "60-64")) %>% 
  select(!c(hb2019, eligible, screened)) %>% 
  mutate(simd = recode(simd, "1" = "1 - Most deprived", "5" = "5 - Least deprived"))

# Table for printing
kpi1_1c_table <- kpi1_1c_scotland %>% 
  pivot_wider(names_from = age_group,
              values_from = age_appr_coverage)  %>% 
  select(!c(hbres_screen, fin_year)) %>% 
  rename("Deprivation (SIMD)" = simd)

# Graph for printing
kpi1_1c_graph1 <- cervical_column_graph(df = kpi1_1c_scotland %>%
                                          select("Age Group" = age_group,
                                                 "Deprivation (SIMD)" = simd,
                                                 "Age-appropriate Coverage (%)" = age_appr_coverage),
                                        data_col = "Age-appropriate Coverage (%)",
                                        colours = create_palette(kpi1_1c_scotland |> 
                                                                   pull(simd)),
                                        threshold = 80)

kpi1_1c_graph2 <- cervical_column_graph(df = kpi1_1c_scotland %>%
                                          select("Deprivation (SIMD)" = simd,
                                                 "Age Group" = age_group,
                                                 "Age-appropriate Coverage (%)" = age_appr_coverage),
                                        data_col = "Age-appropriate Coverage (%)",
                                        colours = create_palette(kpi1_1c_scotland |> 
                                                                   pull(age_group)),
                                        threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_1c1_", year_label, ".png"),
       plot = kpi1_1c_graph1,
       width = 7.5, height = 4)
ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_1c2_", year_label, ".png"),
       plot = kpi1_1c_graph2,
       width = 7.5, height = 4)

## KPI 1.2 - Uptake by HBs, 5 year trend data ----
hist_upt_table <- hist_upt_add %>% 
  filter(age_group == "25-64") %>% 
  select(!c(hb2019, age_group, invites, screened)) 

# Table for printing
kpi1_2_5yr <- hist_upt_table %>% 
  #filter(fin_year %in% fin_year_5) %>% 
  pivot_wider(names_from = fin_year,
              values_from = uptake) %>% 
  rename("Health Board of Residence" = hbres_invite)

# Graph for printing
kpi1_2_graph <- cervical_column_graph(df = hist_upt_table %>%
                                        filter(fin_year %in% fin_year_5) %>%
                                        select("Health Board of Residence" = hbres_invite,
                                               "Financial Year" = fin_year, 
                                               "Uptake (%)" = uptake),
                                      data_col = "Uptake (%)",
                                      colours = create_palette(hist_upt_table |> 
                                                                 filter(fin_year %in% fin_year_5) |> 
                                                                 pull(fin_year)),
                                      threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_2_", year_label, ".png"),
       plot = kpi1_2_graph,
       width = 9, height = 5)

# Scotland Graph for Printing
kpi1_2_graph_scotland <- cervical_line_graph(df = hist_upt_table %>%
                                        filter(hbres_invite=="Scotland") %>%
                                        select("Health Board of Residence" = hbres_invite,
                                               "Financial Year" = fin_year, 
                                               "Uptake (%)" = uptake),
                                      data_col = "Uptake (%)",
                                      colours = create_palette(1),
                                      threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_2_", year_label, "_scotland.png"),
       plot = kpi1_2_graph_scotland,
       width = 9, height = 5)

## KPI 1.2 - Uptake by SIMD, 5 year trend data ----
hist_simd_upt_table <- hist_simd_upt_add %>% 
  filter(age_group == "25-64") %>% 
  select(!c(hb2019, age_group, invites, screened)) %>% 
  mutate(simd = recode(simd, "1" = "1 - Most deprived", "5" = "5 - Least deprived"))

# Tables for printing
kpi1_2a_scotland <- hist_simd_upt_table %>% 
  #filter(fin_year %in% fin_year_5) %>% 
  filter(hbres_invite == "Scotland") 

kpi1_2a_5yr_table <- kpi1_2a_scotland %>% 
  pivot_wider(names_from = fin_year,
              values_from = uptake)  %>% 
  select(!hbres_invite) %>% 
  rename("Deprivation (SIMD)" = simd)

kpi1_2a_1yr_table <- hist_simd_upt_table %>% 
  filter(fin_year == fin_yr) %>% 
  select(!fin_year) %>% 
  pivot_wider(names_from = hbres_invite,
              values_from = uptake) %>% 
  rename("Deprivation (SIMD)" = simd)

# Graph for printing
kpi1_2a_graph <- cervical_column_graph(df = kpi1_2a_scotland %>%
                                         filter(fin_year %in% fin_year_5) %>% 
                                         select("Deprivation (SIMD)" = simd, "Financial Year" = fin_year,
                                                "Uptake (%)" = uptake),
                                       data_col = "Uptake (%)", 
                                       colours = create_palette(kpi1_2a_scotland |> 
                                                                  filter(fin_year %in% fin_year_5) |> 
                                                                  pull(fin_year)),
                                       threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_2a_", year_label, ".png"),
       plot = kpi1_2a_graph,
       width = 7.5, height = 4)

## KPI 1.2 - Coverage by Age group, 5 year trend data ----
hist_upt_age_table <- hist_upt_add %>%  
  filter(age_group %in% c("25-29", "30-34", "35-39", "40-44", 
                          "45-49", "50-54", "55-59", "60-64", "65+")) %>% 
  select(!c(hb2019, invites, screened))

# Table for printing
kpi1_2b_scotland <- hist_upt_age_table %>% 
  #filter(fin_year %in% fin_year_5) %>% 
  filter(hbres_invite == "Scotland")

kpi1_2b_5yr_table <- kpi1_2b_scotland %>% 
  select(!hbres_invite) %>% 
  pivot_wider(names_from = fin_year,
              values_from = uptake) %>% 
  rename("Age Group" = age_group)

kpi1_2b_1yr_table <- hist_upt_age_table %>% 
  filter(fin_year == fin_yr) %>% 
  select(!fin_year) %>% 
  pivot_wider(names_from = hbres_invite,
              values_from = uptake) %>% 
  rename("Age Group" = age_group)

# Graph for printing
kpi1_2b_graph <- cervical_column_graph(df = kpi1_2b_scotland %>%
                                         filter(fin_year %in% fin_year_5) %>% 
                                         select("Age Group" = age_group, 
                                                "Financial Year" = fin_year, 
                                                "Uptake (%)" = uptake),
                                       data_col = "Uptake (%)", 
                                       colours = create_palette(kpi1_2b_scotland |> 
                                                                  filter(fin_year %in% fin_year_5) |> 
                                                                  pull(fin_year)),
                                       threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_2b_", year_label, ".png"),
       plot = kpi1_2b_graph,
       width = 7.5, height = 4)

## KPI 1.2 - Uptake by SIMD and Age group, Scotland level, current financial year ----
kpi1_2c_scotland <- hist_simd_upt_add %>% 
  filter(fin_year == fin_yr,
         hbres_invite == "Scotland") %>% 
  filter(age_group %in% c("25-29", "30-34", "35-39", "40-44", 
                          "45-49", "50-54", "55-59", "60-64", "65+")) %>% 
  select(!c(hb2019, invites, screened)) %>% 
  mutate(simd = recode(simd, "1" = "1 - Most deprived", "5" = "5 - Least deprived"))

# Table for printing
kpi1_2c_table <- kpi1_2c_scotland %>% 
  select(!c(hbres_invite, fin_year)) %>% 
  pivot_wider(names_from = age_group,
              values_from = uptake) %>% 
  rename("Deprivation (SIMD)" = simd)

# Graph for printing
kpi1_2c_graph1 <- cervical_column_graph(df = kpi1_2c_scotland %>%
                                          select("Age Group" = age_group,
                                                 "Deprivation (SIMD)" = simd,
                                                 "Uptake (%)" = uptake),
                                        data_col = "Uptake (%)", 
                                        colours = create_palette(kpi1_2c_scotland |> 
                                                                   pull(simd)),
                                        threshold = 80)

kpi1_2c_graph2 <- cervical_column_graph(df = kpi1_2c_scotland %>%
                                          select("Deprivation (SIMD)" = simd,
                                                 "Age Group" = age_group,
                                                 "Uptake (%)" = uptake),
                                        data_col = "Uptake (%)", 
                                        colours = create_palette(kpi1_2c_scotland |> 
                                                                   pull(age_group)),
                                        threshold = 80)

ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_2c1_", year_label, ".png"),
       plot = kpi1_2c_graph1,
       width = 7.5, height = 4)
ggsave(filename = paste0(fy_folder_path, "graphs/kpi1_2c2_", year_label, ".png"),
       plot = kpi1_2c_graph2,
       width = 7.5, height = 4)


## Write KPI 1 workbook ----
wrt_xlsx_kpi1(scr_period_1yr, scr_period_5yr, fmt_styles, 
              kpi_notes$kpi1_notes, kpi1_list,
              kpi1_1_5yr, kpi1_1_graph,
              kpi1_1a_1yr_table, kpi1_1a_5yr_table, kpi1_1a_graph,
              kpi1_1b_1yr_table, kpi1_1b_5yr_table, kpi1_1b_graph,
              kpi1_1c_table, kpi1_1c_graph1, kpi1_1c_graph2,
              kpi1_2_5yr, kpi1_2_graph,
              kpi1_2a_1yr_table, kpi1_2a_5yr_table, kpi1_2a_graph,
              kpi1_2b_1yr_table, kpi1_2b_5yr_table, kpi1_2b_graph,
              kpi1_2c_table, kpi1_2c_graph1, kpi1_2c_graph2,
              save_to_kpi1_xlsx)


# End of Script ----

