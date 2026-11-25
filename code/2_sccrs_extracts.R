################################################################################
## R Cervical Screening Publication - KPI 1-3
## 2_sccrs_extracts.R
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
pacman::p_load(phsmethods, phslookups, data.table, readr, dplyr, 
               tidyr, tidylog, arrow, stringr, janitor, lubridate, 
               writexl, openxlsx)

# Clean global environment
rm(list = ls())

setwd("/PHI_conf/CancerGroup1/Topics/Git/John/cervical_kpi_dev_uat/cervical_kpi_dev_uat")


# Source variables
source("code/1_housekeeping.R")
source("src/func_cervical_kpi_calc_graph.R")

# Clean global environment
rm(list = ls(pattern = "fig|kpi|fmt|scr|hb|graph|pct|no_data"))


# Exclusion types:
flagged_exclusions <- c(
  "Deceased", "Deleted", "GANA", "Institutionalised",
  "Marked For Deletion", "Military Service", "No Cervix",
  "No Further Recall", "Opted Out", "Redundant", "TR",
  "Terminally Ill", "Transferred Out By SCCRS",
  "Transferred Out Untraced", "Transferred Out by CHI",
  "Transferred Out of Scotland"
)

# unflagged_exclusions not used, but listed for transparency

unflagged_exclusions <- c(
  "Anatomically Impossible", "At Colposcopy", "Co-Morbidity",
  "Defaulter", "Learning Disabilities",
  "New Registration in Progress", "Not Clinically Appropriate",
  "Pregnant", "Smear In Progress", "Smear at Colposcopy",
  "Student Deferment"
)


# 2: Read in data ----
## 2.1: cervical files from ATOS ----
cervical_demographics <- read_parquet(
  paste0(data_file_path, "SCCRS_KPI_Demographics_", extract_dl_date, ".parquet")
) %>%
  mutate(dob = dmy(dob)) %>%
  mutate(registered_date = dmy(registered_date)) %>%
  mutate(
    pc8 = format_postcode(current_postcode, format = "pc8"),
    .after = current_postcode
  )

cervical_exclusions <- read_parquet(
  paste0(data_file_path, "SCCRS_KPI_Exclusion_", extract_dl_date, ".parquet")
) %>%
  mutate(
    exclusion_start_date = dmy(exclusion_open_date),
    .before = exclusion_open_date
  ) %>% # rename to match SCCRS database
  mutate(exclusion_open_date = dmy(dateof_removafrom_call_recall)) %>% # rename to match SCCRS database
  select(!dateof_removafrom_call_recall) %>% # Remove char column
  mutate(exclusion_end_date = dmy(exclusion_end_date))

cervical_results <- read_parquet(
  paste0(data_file_path, "SCCRS_KPI_Result_", extract_dl_date, ".parquet")
) %>%
  mutate(dateof_examination = dmy(dateof_examination)) %>%
  mutate(rejectedby_lab_date = dmy(rejectedby_lab_date)) %>%
  mutate(cytology_bookedin_date = dmy(cytology_bookedin_date)) %>%
  mutate(virology_bookedin_date = dmy(virology_bookedin_date)) %>%
  mutate(date_lab_hpv_result_reported = dmy(date_lab_hpv_result_reported)) %>%
  mutate(date_lab_cytology_result_reported = dmy(date_lab_cytology_result_reported)) %>%
  mutate(authorised_date = date(dmy_hms(authorised_date))) %>%
  mutate(referral_date = date(authorised_date) + 1) %>%
  mutate(result_number = as.double(result_number)) %>%
  mutate(pc8 = format_postcode(postcode, format = "pc8"), .after = postcode)

cervical_prompts <- read_parquet(
  paste0(data_file_path, "SCCRS_KPI_Prompt_", extract_dl_date, ".parquet")
) %>%
  mutate(dateof_prompt = dmy(dateof_prompt)) %>%
  mutate(pc8 = format_postcode(postcode, format = "pc8"), .after = postcode)

gc()
## 2.2: Read in simd data ----
simd <- get_simd_postcode() %>%
  select(pc8, simd2020v2_sc_quintile, hb2019, hb2019name)

# 3: Reformat Data: -----

## 3.1: Demographics Extract ----
cervical_demographics_single <- cervical_demographics |>
  distinct() |>
  rename(pc8_demo = pc8) |>
  filter(registered_date <= dmy(fy_end_date))

## 3.2: Exclusion extract ----
# Exclusion rules:
# Exclusions are not included if start date is after end of financial year AND
# either no end date (NA) OR end date after end of financial year.
cervical_exclusions_open <- cervical_exclusions %>%
  filter(type_of_exclusion %in% flagged_exclusions) |>
  select(upi_number,
         practice_exc = practice,
         type_of_exclusion,
         exclusion_start_date,
         exclusion_open_date,
         exclusion_end_date,
         exclusion_close_reason
  ) |>
  mutate(
    exclusion_end_date = coalesce(
      exclusion_end_date,
      # pick far off date to if there is no date to allow merging below
      ymd("9999-12-31") 
    )
  ) |>
  arrange(upi_number, exclusion_start_date, exclusion_end_date)

## 3.3: Prompts Extract ----
# Reformat prompts data to flag which invites appear in each financial year
cervical_prompt_data <- cervical_prompts %>%
  filter(dateof_prompt >= fy_start_date_5) |>
  select(upi_number, prompt_type, screeningstatus, dateof_prompt, pc8_invite = pc8, practice_invite = practice) %>%
  arrange(upi_number, desc(dateof_prompt)) |>
  mutate(
    financial_year = create_fin_year(dateof_prompt)
  ) |>
  filter(financial_year %in% fin_year_5) 

## 3.4: Results Extract ----
cervical_results_tidy <- cervical_results |>
  rename(pc8_screen = pc8, practice_res = practice) |>
  distinct() |>
  arrange(upi_number, dateof_examination)

rm(cervical_demographics, cervical_prompts, cervical_results, cervical_exclusions)

gc()

# 4: Create KPI extracts ----
## 4.1: KPI 1 ----
### 4.1.1: KPI 1.1 Extract ----
# Assign financial years to upi records
cervical_demographics_fin_yr <- cervical_demographics_single |>
  crossing(fin_year_5) |>
  rename(financial_year = fin_year_5) |>
  mutate(
    start_year = as.integer(str_sub(financial_year, 1, 4)),
    year_start = make_date(start_year, 4, 1),
    year_end = make_date(start_year + 1L, 3, 31)
  ) |>
  filter(registered_date <= year_end) |>
  add_age() |>
  mutate(
    start_date_5_5 = make_date(year(year_start) - 5, 10, 1),
    start_date_3_5 = make_date(year(year_start) - 3, 10, 1),
  )

gc()


# Join demographics data with exclusions and results data. Prompts data is also 
# joined for assigning postcode. Then join with Deprivation data
kpi1_1_extract <- cervical_demographics_fin_yr |>
  left_join(cervical_exclusions_open,
            by = join_by(
              upi_number,
              year_end >= exclusion_open_date,
              year_end < exclusion_end_date
            ),
            multiple = "last"
  ) |>
  mutate(exclusion_end_date = na_if(exclusion_end_date, ymd("9999-12-31"))) |> 
  left_join(cervical_results_tidy,
            by = join_by(
              upi_number,
              year_end >= dateof_examination
            ),
            multiple = "last"
  ) |>
  mutate(examination_minus_6m = dateof_examination - days(183)) |> 
  left_join(select(cervical_prompt_data, upi_number, dateof_prompt, pc8_invite, practice_invite),
            by = join_by(upi_number, 
                         dateof_examination >= dateof_prompt,
                         examination_minus_6m <= dateof_prompt
            ),
            multiple = "last") |>
  select(!examination_minus_6m) |> 
  mutate(
    scr5_5 = if_else(!is.na(dateof_examination) &
                       dateof_examination >= start_date_5_5, 1, 0),
    scr3_5 = if_else(age <= 49 & 
                       !is.na(dateof_examination) &
                       dateof_examination >= start_date_3_5, 1, 0)
  ) |>
  mutate(screen = if_else(financial_year < "2025/26" & age <= 49, 
                          scr3_5, scr5_5)) %>%
  add_flags() |>
  # Eligible lower age changes based on financial year
  mutate(eligible = case_when(
    exclusion_flag == 1 ~ 0,
    financial_year <= "2015/16" & age >= 20 & age <= 60 ~ 1,
    financial_year == "2016/17" & age >= 21 & age <= 64 ~ 1,
    financial_year == "2017/18" & age >= 22 & age <= 64 ~ 1,
    financial_year == "2018/19" & age >= 23 & age <= 64 ~ 1,
    financial_year == "2019/20" & age >= 24 & age <= 64 ~ 1,
    financial_year >= "2020/21" & age >= 25 & age <= 64 ~ 1,
    .default = 0
  )) |>
  filter(eligible == 1) |> 
  add_postcode() |>
  select(financial_year,
         "NHS Board of Residence" = hb2019name,
         "Deprivation Category" = simd2020v2_sc_quintile,
         "CHI" = upi_number,
         dob, age, eligible, scr5_5, scr3_5, screen, everything()
  )


table(kpi1_1_extract$financial_year)

rm(cervical_demographics_fin_yr)

gc()

### 4.1.2: KPI 1.2 Extract ----

# Join demographics data with prompt and results data. Then join with 
# Deprivation data
kpi1_2_extract <- cervical_prompt_data |>
  mutate(
    invite_6m = dateof_prompt + days(183),
    start_year = as.integer(str_sub(financial_year, 1, 4)),
    year_start = make_date(start_year, 4, 1),
    year_end = make_date(start_year + 1L, 3, 31)
  ) |>
  left_join(select(cervical_results_tidy, upi_number, practice_res, pc8_screen, 
                   dateof_examination),
            by = join_by(
              upi_number,
              dateof_prompt <= dateof_examination,
              invite_6m >= dateof_examination
            ),
            multiple = "last"
  ) |> 
  group_by(upi_number, financial_year) |> 
  # filter to keep most recent prompt match of dateof_examination result for 
  # each upi/year
  filter(row_number() == if (all(is.na(dateof_examination))) { 
    1L } else { 
      which.max(dateof_examination) } ) |> 
  ungroup() |>
  left_join(cervical_demographics_single,
            by = join_by(
              upi_number,
              year_end >= registered_date
            )) |>
  left_join(cervical_exclusions_open,
            by = join_by(
              upi_number,
              year_end >= exclusion_open_date,
              year_end < exclusion_end_date
            ),
            multiple = "last") |>
  mutate(exclusion_end_date = na_if(exclusion_end_date, ymd("9999-12-31"))) |> 
  add_age() |>
  add_flags() |>
  add_postcode() |>
  select(upi_number,
         "NHS Board of Residence" = hb2019name,
         "Deprivation Category" = simd2020v2_sc_quintile,
         dob, uptake_flag, everything()
  )

gc()

## 4.2: KPI 2 Extract ----
# Flag or results either arriving in financial year or reported in financial 
# year. 
cervical_results_kpi2 <- cervical_results_tidy |>
  filter(authorised_date >= fy_start_date_5 |
           rejectedby_lab_date >= fy_start_date_5 |
           virology_bookedin_date >= fy_start_date_5 |
           cytology_bookedin_date >= fy_start_date_5) |>
  mutate(
    arrived_date = coalesce(
      rejectedby_lab_date,
      virology_bookedin_date,
      cytology_bookedin_date
    ),
    reported_in_fy = if_else(is.na(authorised_date),
                             NA_character_,
                             paste0(
                               year(authorised_date %m-% months(3)), "/",
                               substr(year(authorised_date %m-% months(3))+ 1, 3, 4)
                             )
    ),
    arrived_in_fy = if_else(is.na(arrived_date),
                            NA_character_,
                            paste0(
                              year(arrived_date %m-% months(3)), "/",
                              substr(year(arrived_date %m-% months(3)) + 1,3, 4)
                            )
    )
  ) |>
  filter(reported_in_fy %chin% fin_year_5 | arrived_in_fy %chin% fin_year_5) |>
  mutate(
    reported_fy = if_else(!is.na(reported_in_fy), 1, 0),
    arrived_fy = if_else(!is.na(arrived_in_fy), 1, 0)
  ) |>
  pivot_longer(
    cols = c(reported_in_fy, arrived_in_fy),
    names_to = "fy_source",
    values_to = "financial_year"
  ) |>
  filter(!is.na(financial_year) & financial_year %in% fin_year_5) %>%
  distinct(upi_number, financial_year, .keep_all = TRUE) |>
  select(-fy_source) |>
  mutate(
    # Add flag for results reported within 14 calendar days after screen
    reported_14 = if_else(
      between(
        as.integer(authorised_date - dateof_examination), 0, 14
      ),
      1, 0, missing = 0
    ),
    # Add flag for any samples rejected by the lab
    rejected = if_else(sample_rejectedby_lab == "Y", 1, 0)
  )

# Join results extract to demographics and prompts data. Propmt data only used
# to assign postcode
kpi2_extract <- cervical_results_kpi2 |>
  left_join(cervical_demographics_single, by = "upi_number") |>
  mutate(examination_minus_6m = dateof_examination - days(183)) |> 
  left_join(select(cervical_prompt_data, upi_number, dateof_prompt, pc8_invite, practice_invite),
            by = join_by(upi_number, 
                         dateof_examination >= dateof_prompt,
                         examination_minus_6m <= dateof_prompt
            ),
            multiple = "last") |>
  mutate(
    start_year = as.integer(str_sub(financial_year, 1, 4)),
    year_start = make_date(start_year, 4, 1),
    year_end = make_date(start_year + 1L, 3, 31)
  ) |>
  add_age() |>
  add_flags(exclusion = FALSE) |>
  add_postcode() |>
  select(!c(start_year, year_start, year_end, examination_minus_6m)) |>
  select(financial_year,
         "NHS Board of Residence" = hb2019name,
         "Deprivation Category" = simd2020v2_sc_quintile,
         dob, age, reported_fy, reported_14, arrived_fy, rejected, everything()
  )

rm(cervical_results_kpi2)
gc()

## 4.3: KPI 3 Extract ----
# Create a look up for financial years
fy_lookup <- tibble(financial_year = fin_year_5) %>%
  mutate(
    start_year = as.integer(str_sub(financial_year, 1, 4)),
    year_start = make_date(start_year, 4, 1),
    year_end = make_date(start_year + 1, 3, 31),
    year_end_6m = make_date(start_year + 1, 9, 30)
  ) |>
  select(!start_year)

# Identify samples reported within the financial year with cytology and HPV 
# results as well as samples reported within 6 months of their screening date
cervical_results_kpi3 <- cervical_results_tidy %>%
  filter(authorised_date >= fy_start_date_5) |>
  left_join(fy_lookup,
            by = join_by(between(authorised_date, year_start, year_end_6m)),
            multiple = "all") |>
  filter(financial_year %in% fin_year_5) |>
  mutate(
    reported_fy = if_else(between(authorised_date, year_start, year_end),
                          1, 0, missing = 0),
    reported_fy_6m = if_else(between(authorised_date, year_start, year_end_6m),
                             1, 0, missing = 0)
  ) |>
  mutate(
    reported_cyt_fy = if_else(reported_fy == 1 & !cytology_test_result %in% c(""),
                              1, 0, missing = 0),
    reported_cyt_unsat = if_else(cytology_test_result == "Unsatisfactory",
                                 1, 0, missing = 0),
    reported_hpv_fy = if_else(reported_fy == 1 & !hpv_rest_result %in% c(""),
                              1, 0, missing = 0),
    reported_hpv_fail = if_else(hpv_rest_result == "hr-HPV Fail",
                                1, 0, missing = 0),
    # Add flag for when a result was reported as Fail or Unsatisfactory within 
    # the financial year
    reported_fail_unsat = if_else(
      reported_fy == 1 &
        (hpv_rest_result == "hr-HPV Fail" &
           (cytology_test_result %in% c("","Negative","Unsatisfactory")) |
           cytology_test_result == "Unsatisfactory"),
      1, 0, missing = 0
      )
  ) |> 
  filter(reported_cyt_fy == 1 | reported_hpv_fy == 1 | reported_fy == 1 | 
           reported_fy_6m == 1)

gc()

### 4.3.1: KPI 3.1 and KPI 3.2 ----
# Join results with demographics and prompt data. Prompt data is only used 
# for assigning postcode.
kpi3_extract <- cervical_results_kpi3 |> 
  left_join(cervical_demographics_single,
            by = join_by(upi_number)) |> 
  mutate(examination_minus_6m = dateof_examination - days(183)) |> 
  left_join(select(cervical_prompt_data, upi_number, dateof_prompt, pc8_invite, practice_invite),
            by = join_by(upi_number, 
                         dateof_examination >= dateof_prompt,
                         examination_minus_6m <= dateof_prompt
            ),
            multiple = "last") |>
  add_age() |> 
  add_flags(exclusion = FALSE) |> 
  add_postcode() |> 
  select(!c(pc8_invite, practice_invite, 
            year_start, year_end, year_end_6m,
            invite_fy_flag, uptake_flag,
            examination_minus_6m)) |> 
  select(financial_year,
         upi_number,
         "NHS Board of Residence" = hb2019name, 
         "Deprivation Category" = simd2020v2_sc_quintile,
         dob, age, reported_cyt_fy, reported_cyt_unsat, reported_hpv_fy, 
         reported_hpv_fail, everything())

gc()

### 4.3.2: KPI 3.3 Extract ----

# Total number of individuals with an unsatisfacotry result in financial year
test_fail_upis <- cervical_results_kpi3 %>%
  filter(reported_fy == 1 & reported_fail_unsat == 1) %>% 
  select(upi_number, financial_year) |> 
  mutate(failed_test_fy = 1) |> 
  distinct_all()

cervical_results_kpi3_3  <- cervical_results_kpi3 %>% 
  # Any result reported within FY + 6 months (to include repeat screen for end 
  # of year failed tests)
  filter(reported_fy_6m == 1) %>% 
  # Anyone with a failed test within the FY (including additional records before 
  # or after failed test result)
  inner_join(test_fail_upis, by = join_by(upi_number, financial_year)) %>% 
  filter(failed_test_fy == 1) |> 
  select(!failed_test_fy) |> 
  mutate(first_reported_fy = reported_fy_6m) %>% 
  # sort by result number to get data in chronological order
  arrange(financial_year, result_number) %>% 
  group_by(financial_year, upi_number) %>% 
  # Add new column with numbered visit for each individual
  mutate(visit = 1:n()) |> 
  # flag first fail a person has each financial year
  mutate(first_fail = case_when(reported_fail_unsat == 1 ~ visit),
         first_fail = case_when(is.na(first_fail) ~ max(first_fail, na.rm = TRUE),
                                .default = first_fail)) %>% 
  filter(visit == max(visit)) %>% 
  ungroup() %>% 
  mutate(repeat_scr = ifelse(first_fail < visit, 1, 0)) 

gc()

# Join results data with demographic and prompts data. Prompt data is used for
# assigning postcode.
kpi3_3_extract <- cervical_results_kpi3_3 |> 
  left_join(cervical_demographics_single,
            by = join_by(upi_number)) |> 
  mutate(examination_minus_6m = dateof_examination - days(183)) |> 
  left_join(select(cervical_prompt_data, upi_number, dateof_prompt, pc8_invite, practice_invite),
            by = join_by(upi_number, 
                         dateof_examination >= dateof_prompt,
                         examination_minus_6m <= dateof_prompt
            ),
            multiple = "last") |>
  add_age() |> 
  add_flags(exclusion = FALSE) |> 
  add_postcode() |> 
  select(!c(pc8_invite, practice_invite, 
            year_start, year_end, year_end_6m,
            invite_fy_flag, uptake_flag,
            examination_minus_6m)) |> 
  select(financial_year,
         upi_number,
         "NHS Board of Residence" = hb2019name, 
         "Deprivation Category" = simd2020v2_sc_quintile,
         dob, age, first_reported_fy, repeat_scr, everything())

rm(cervical_results_kpi3, cervical_results_kpi3_3)
gc()

# 5: Save outputs ----

write_parquet(
  kpi1_1_extract,
  paste0(fy_folder_path, "Checks/", "kpi1_1_extract_update.parquet")
)
write_parquet(
  kpi1_2_extract,
  paste0(fy_folder_path, "Checks/", "kpi1_2_extract_update.parquet")
)
write_parquet(
  kpi2_extract,
  paste0(fy_folder_path, "Checks/", "kpi2_extract_update.parquet")
)
write_parquet(
  kpi3_extract,
  paste0(fy_folder_path, "Checks/", "kpi3_extract_update.parquet")
)
write_parquet(
  kpi3_3_extract,
  paste0(fy_folder_path, "Checks/", "kpi3_3_extract_update.parquet")
)

gc()

#### End of Script ####
