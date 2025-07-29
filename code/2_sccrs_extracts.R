##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## R Cervical Screening - KPIs for Publication
## 2_sccrs_extracts.R
## 
## Adult Screening Team
## Created: 27/10/24
## Last update: 09/07/2025
## 
## Written/run on RStudio Posit Workbench
## Platform: x86_64-pc-linux-gnu
## R version 4.4.2 (Pile of Leaves) - 2024-10-31
##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Packages ----
library(data.table)
library(readr)
library(dplyr)
library(tidyr)
library(tidylog)
library(janitor)
library(lubridate)
library(phsmethods)

## Clean global environment
rm(list=ls())

## Sourcing
source("code/1_housekeeping.R")
source("src/func_cervical_kpi_calc_graph.R")

## Clean global environment
rm(list = ls(pattern = "fig|kpi|fmt|scr|hb|graph|pct|no_data"))


## Exclusion types
flagged_exclusions <- c("Deceased", "Deleted", "GANA", "Institutionalised", 
                        "Marked For Deletion", "Military Service", "No Cervix", 
                        "No Further Recall", "Opted Out", "Redundant", "TR", 
                        "Terminally Ill", "Transferred Out By SCCRS", 
                        "Transferred Out Untraced", "Transferred Out by CHI", 
                        "Transferred Out of Scotland")

unflagged_exclusions <- c("Anatomically Impossible", "At Colposcopy", "Co-Morbidity", 
                          "Defaulter", "Learning Disabilities", 
                          "New Registration in Progress", "Not Clinically Appropriate",
                          "Pregnant", "Smear In Progress", "Smear at Colposcopy",
                          "Student Deferment")
### unflagged_exclusions not used, but listed for transparency


# Read in SCCRS extracts ----
cervical_demographics <- fread(paste0(fy_folder_path, "data/SCCRS_KPI_Demographics_",
                                      extract_dl_date, ".csv"),
                               colClasses = "character") %>% 
  clean_names() 

cervical_exclusions <- fread(paste0(fy_folder_path, "data/SCCRS_KPI_Exclusion_",
                                    extract_dl_date, ".csv"),
                             colClasses = "character") %>% 
  clean_names()

cervical_results <- fread(paste0(fy_folder_path, "data/SCCRS_KPI_Result_",
                                 extract_dl_date, ".csv"),
                          colClasses = "character") %>%
  clean_names() 

cervical_prompts <- fread(paste0(fy_folder_path, "data/SCCRS_KPI_Prompt_",
                                 extract_dl_date, ".csv"),
                          colClasses = "character") %>%
  clean_names()

## Read in SIMD look up ----
simd <- read_rds(simd20_path) %>%
  select(pc8, simd2020v2_sc_quintile, hb2019, hb2019name) 


## For each extract: Check all records have a UPI
cervical_demographics %>% count(is.na(upi_number))
cervical_exclusions %>% count(is.na(upi_number))
cervical_results %>% count(is.na(upi_number))
cervical_prompts %>% count(is.na(upi_number))
##!! If any extract have any records with no UPI, further investigation is required before continuing


# Reformat date and postcode columns ----
cervical_demographics <- cervical_demographics %>% 
  mutate(dob = dmy(dob)) %>%
  mutate(registered_date = dmy(registered_date)) %>% 
  mutate(pc8 = format_postcode(current_postcode, format = "pc8"), .after = current_postcode)

cervical_prompts <- cervical_prompts %>% 
  mutate(dateof_prompt = dmy(dateof_prompt)) %>% 
  mutate(pc8 = format_postcode(postcode, format = "pc8"), .after = postcode)

cervical_exclusions <- cervical_exclusions %>% 
  mutate(exclusion_start_date = dmy(exclusion_open_date), .before = exclusion_open_date) %>% # rename to match SCCRS database
  mutate(exclusion_open_date = dmy(dateof_removafrom_call_recall)) %>% # rename to match SCCRS database
  select(!dateof_removafrom_call_recall) %>% # Remove char column
  mutate(exclusion_end_date = dmy(exclusion_end_date))

cervical_results <- cervical_results %>% 
  mutate(dateof_examination = dmy(dateof_examination)) %>% 
  mutate(rejectedby_lab_date = dmy(rejectedby_lab_date)) %>% 
  mutate(cytology_bookedin_date = dmy(cytology_bookedin_date)) %>%
  mutate(virology_bookedin_date = dmy(virology_bookedin_date)) %>% 
  mutate(date_lab_hpv_result_reported = dmy(date_lab_hpv_result_reported)) %>% 
  mutate(date_lab_cytology_result_reported = dmy(date_lab_cytology_result_reported)) %>% 
  mutate(authorised_date = dmy_hms(authorised_date)) %>%
  mutate(referral_date = date(authorised_date)+1) %>% 
  mutate(result_number = as.double(result_number)) %>% 
  mutate(pc8 = format_postcode(postcode, format = "pc8"), .after = postcode)


# Prepare data frame for KPI 1 ----
## Demographics extract, prepared to one record per CHI
## Prompts extract, any prompt within FY (KPI 1.2)
## Exclusions extract, any open exclusion at time of end of FY, one per CHI
## Results extract, most recent screening date before end of FY, one record per CHI
## SIMD look up
## 
# Prepare data frame for KPI 2, KPI 3 ----
## Use Demographics extract as prepared for KPI 1
## Result extract, with flags appropriate for each KPI
## SIMD look up


# Demographics Extract ----
## Calculate age at time of financial year end date and create age groups
cervical_demographics_dob <- cervical_demographics %>%
  filter(registered_date <= dmy(fy_end_date)) %>% 
  mutate(age = age_calculate(dob, dmy(fy_end_date)+1)) %>% 
  mutate(age_group = case_when(age <= 24 ~ "Under 25",
                               age >= 25 & age <=29 ~ "25-29",
                               age >= 30 & age <=34 ~ "30-34",
                               age >= 35 & age <=39 ~ "35-39",
                               age >= 40 & age <=44 ~ "40-44",
                               age >= 45 & age <=49 ~ "45-49",
                               age >= 50 & age <=54 ~ "50-54",
                               age >= 55 & age <=59 ~ "55-59",
                               age >= 60 & age <=64 ~ "60-64",
                               age >= 65 ~ "65+"),
         age_group_25_64 = case_when(age >= 25 & age <= 64 ~ "25-64",
                                     age <= 24 ~ "Under 25", 
                                     age >= 65 ~ "65+"),
         age_group2549_5064 = case_when(age <= 24 ~ "Under 25",
                                        age >= 25 & age <= 49 ~ "25-49",
                                        age >= 50 & age <= 64 ~ "50-64",
                                        age >= 65 ~ "65+"))

## Number of unique UPIs
length(cervical_demographics_dob$upi_number)
length(unique(cervical_demographics_dob$upi_number))
##!! Check number of total records are equal to number of unique records

rm(cervical_demographics)
## If rerunning multiple years, do not remove dataframe
## Reuse and rerun financial year by updating/sourcing HK script, and run from Section: Demographics Extract


# Prompt extract ----
## Select prompts sent within the financial year 
cervical_prompts_fy <- cervical_prompts %>% 
  filter(dateof_prompt <= dmy(fy_end_date),
         dateof_prompt >= dmy(fy_start_date))

## Pivot wider type of prompt and associated dates to get one row per UPI
cervical_prompt_data <- cervical_prompts_fy %>% 
  select(upi_number, prompt_type, dateof_prompt, pc8) %>% 
  pivot_wider(id_cols = upi_number,
              names_from = prompt_type,
              values_from = c(dateof_prompt, pc8),
              values_fn = max) %>%
  mutate(pc8_R1 = coalesce(pc8_R1, pc8_I),
         pc8_R2 = coalesce(pc8_R2, pc8_R1),
         pc8_invite = coalesce(pc8_R3, pc8_R2)) %>% 
  select(upi_number, pc8_invite, 
         I = dateof_prompt_I, R1 = dateof_prompt_R1, 
         R2 = dateof_prompt_R2, R3 = dateof_prompt_R3)

## Number of unique UPIs
length(cervical_prompt_data$upi_number)
length(unique(cervical_prompt_data$upi_number))
##!! Check number of total records are equal to number of unique records


rm(cervical_prompts, cervical_prompts_fy)
## If rerunning multiple years, do not remove dataframe
## Reuse and rerun financial year by updating/sourcing HK script, and run from Section: Prompt extract


# Exclusion extract ----
## Exclusion rules:
### Exclusions are deemed Open if open date is before end of financial year AND
### either no end date (NA) OR end date after end of financial year.
cervical_exclusions_open <- cervical_exclusions %>% 
  filter(exclusion_open_date <= dmy(fy_end_date), # <= including end of FY, exclusion open date on or before 31/03/YY
         is.na(exclusion_end_date) | exclusion_end_date > dmy(fy_end_date)) %>% # exclusion end date after 31/03/YY or NA
  filter(type_of_exclusion %in% flagged_exclusions)

## Identify latest exclusion where there are >1 open exclusion records per chi
### Convert to data.table
setDT(cervical_exclusions_open)

### Find row index with first exclusion for each upi
latest_exclusion_indexes <- cervical_exclusions_open[,.(I=.I[which.max(exclusion_start_date)]), keyby=upi_number]$I

### Use index rows to filter data then select relevant columns
latest_exclusion  <- cervical_exclusions_open[latest_exclusion_indexes, .(upi_number,
                                                         type_of_exclusion,
                                                         exclusion_start_date,
                                                         exclusion_open_date,
                                                         exclusion_end_date,
                                                         exclusion_close_reason
                                                        )
                                              ]

# Convert back to tibble 
latest_exclusion <- as_tibble(latest_exclusion)  

# Number of unique UPIs
length(latest_exclusion$upi_number)
length(unique(latest_exclusion$upi_number))
##!! Check number of total records are equal to number of unique records

rm(cervical_exclusions, cervical_exclusions_open,latest_exclusion_indexes)
## If rerunning multiple years, do not remove dataframe
## Reuse and rerun financial year by updating/sourcing HK script, and run from Section: Exclusion extract


# Result extract ----
## Identify the most recent result recorded per chi using the highest Result Number
## Checking for any NAs in the result number column and their hpv/cytology result
cervical_results %>% 
  filter(is.na(result_number)) %>% 
  count(hpv_rest_result, cytology_test_result)
##!! Check if any reported results are missing a result number, further investigation required if identified

## Create a data.table version of cervical results
cervical_results_dt <- as.data.table(cervical_results)

## KPI 1.1, Identify most recent screen for each UPI before end of financial year ----
cervical_results_kpi1_1 <-cervical_results_dt[
  # Remove records where any screen performed after the end of the financial year
  dateof_examination <= dmy(fy_end_date),
  # drop columns that's within the demographics data to remove duplicate potentially caused by these variables
  .(upi_number, pc8, dateof_examination,result_number) ][
      # Sort data before removing duplicate rows
      order(upi_number, -dateof_examination, -result_number) 
]

### Find row index with max date of examination for each upi
cervical_results_kpi1_1_indexes <- cervical_results_kpi1_1[,.(I=.I[which.max(dateof_examination)]), keyby=upi_number]$I

### Filter rows using index above
cervical_results_kpi1_1 <- as_tibble(cervical_results_kpi1_1[cervical_results_kpi1_1_indexes, 
                                                                   .(upi_number,
                                                                     pc8_screen=pc8,
                                                                     dateof_examination
                                                                    )
                                                                   ]
                                        )

## Number of unique UPIs
length(cervical_results_kpi1_1$upi_number)
length(unique(cervical_results_kpi1_1$upi_number))
##!! Check number of total records are equal to number of unique records

rm(cervical_results_kpi1_1_indexes)
## If rerunning multiple years, do not remove dataframe
## Reuse and rerun financial year by updating/sourcing HK script, and run from Section: Result extract

## KPI 1.2, Identify first screen for each UPI within financial year + 6 months ----
### Using the first screen date, not the most recent as this might be a follow up screen

### Find first and last date of examination
cervical_results_kpi1_2_pw <- cervical_results_dt[,.(upi_number, dateof_examination)][
  # Remove records where any screen performed out with the financial year + 6 months
  dateof_examination <= dmy(fy_end_date_plus_6m) & dateof_examination >= dmy(fy_start_date)][
    # order by date and flag examination number, and find first and last doe (date of examination) by upi
    order(dateof_examination),`:=`(
      visit = paste0("doe_",1:.N),
      first_doe = min(dateof_examination),
      last_doe = max(dateof_examination)
    ), by=upi_number]

### Make table wider
cervical_results_kpi1_2 <- dcast(cervical_results_kpi1_2_pw, upi_number + first_doe + last_doe ~ visit, value.var="dateof_examination")

### move columns "first_doe" and "last_doe" to end for consistency  
setcolorder(cervical_results_kpi1_2, c(setdiff(names(cervical_results_kpi1_2), c("first_doe", "last_doe")), "first_doe", "last_doe"))

### Convert back to tibble for dplyr code
cervical_results_kpi1_2 <- as_tibble(cervical_results_kpi1_2)

## Number of unique UPIs
length(cervical_results_kpi1_2$upi_number)
length(unique(cervical_results_kpi1_2$upi_number))
##!! Check number of total records are equal to number of unique records


## KPI 2, Identify reported results and tests arrived within financial year ----
### Find row index of for results reported within the financial year
reported_in_fy <- cervical_results_dt[, .I[!is.na(authorised_date) & between(as.Date(authorised_date), dmy(fy_start_date), dmy(fy_end_date))]]

### Find row index for samples arrived at the lab within the financial year
arrived_in_fy <- cervical_results_dt[,    .I[ifelse(!is.na(rejectedby_lab_date) == TRUE, 
                                                    between(rejectedby_lab_date, dmy(fy_start_date), dmy(fy_end_date)),
                                                    ifelse(!is.na(virology_bookedin_date) == TRUE,
                                                           between(virology_bookedin_date, dmy(fy_start_date), dmy(fy_end_date)),
                                                           between(cytology_bookedin_date, dmy(fy_start_date), dmy(fy_end_date)))
                                                    )==TRUE]]

### Combine row indexes together
in_fy <- union(arrived_in_fy,reported_in_fy)

### Filter data using row indexes then add flags
cervical_results_kpi2 <- cervical_results_dt[in_fy
  ][,`:=`(
  # Add flag for results reported within the financial year
  reported_fy = case_when(between(as.Date(authorised_date), dmy(fy_start_date), dmy(fy_end_date)) == TRUE ~ 1,
                          .default = 0),
  # Add flag for results reported within 14 calendar days after screen
  reported_14 = case_when(between(interval(dateof_examination, as.Date(stringr::str_sub(authorised_date, 1, 10))) %/% days(1), 0, 14) ~ 1,
                          .default = 0),
  # Add flag for samples arrived at the lab within the financial year
  arrived_fy = case_when(
    ifelse(!is.na(rejectedby_lab_date) == TRUE,
           between(rejectedby_lab_date, dmy(fy_start_date), dmy(fy_end_date)),
           ifelse(!is.na(virology_bookedin_date) == TRUE,
                  between(virology_bookedin_date, dmy(fy_start_date), dmy(fy_end_date)),
                  between(cytology_bookedin_date, dmy(fy_start_date), dmy(fy_end_date)))
    ) ~ 1,
    .default = 0),
  # Add flag for any samples rejected by the lab
  rejected = case_when(sample_rejectedby_lab == "Y" ~ 1,.default = 0)
)]

setnames(cervical_results_kpi2, "pc8", "pc8_result")

cervical_results_kpi2 <- as_tibble(cervical_results_kpi2)

rm(arrived_in_fy,reported_in_fy,in_fy)
## If rerunning multiple years, do not remove dataframe
## Reuse and rerun financial year by updating/sourcing HK script, and run from Section: Result extract


## KPI 3, Identify samples reported within the financial year with cytology and HPV results ----
## as well as samples reported within 6 months of their screening date
cervical_results_kpi3 <- cervical_results %>% 
  # Add flags for results reported within the financial year, if cytology results were unsatisfactory or HPV results were fail,
  # and results reported within 6 months of their screen
  mutate(reported_cyt_fy = case_when(between(as.Date(authorised_date), dmy(fy_start_date), dmy(fy_end_date)) == TRUE &
                                       !cytology_test_result %in% c("") == TRUE ~ 1,
                                     .default = 0),
         reported_cyt_unsat = ifelse(cytology_test_result == "Unsatisfactory", 1, 0)) %>% 
  mutate(reported_hpv_fy = case_when(between(as.Date(authorised_date), dmy(fy_start_date), dmy(fy_end_date)) == TRUE &
                                       !hpv_rest_result %in% c("") == TRUE ~ 1,
                                     .default = 0),
         reported_hpv_fail = ifelse(hpv_rest_result == "hr-HPV Fail", 1, 0)) %>% 
  # Add flags for results reported within the financial year or an additional 6month for follow up screens
  # Add flag for when a result was reported as Fail or Unsatisfactory within the financial year
  mutate(reported_fy = case_when(between(as.Date(authorised_date), dmy(fy_start_date), dmy(fy_end_date)) == TRUE ~ 1, 
                                 .default = 0),
         reported_fy_6m = case_when(between(as.Date(authorised_date), dmy(fy_start_date), dmy(fy_end_date_plus_6m)) == TRUE ~ 1,
                                    .default = 0)) %>%
  mutate(reported_fail_unsat = case_when(reported_fy == 1 & ((hpv_rest_result == "hr-HPV Fail" &
                                                                (cytology_test_result == "" |
                                                                   cytology_test_result == "Negative" |
                                                                   cytology_test_result == "Unsatisfactory")) |
                                                               cytology_test_result == "Unsatisfactory") ~ 1,
                                         .default = 0)) %>% 
  # Remove records where any results were reported outwith the financial year (+ 6 months)
  filter(reported_cyt_fy == 1 | reported_hpv_fy == 1 | reported_fy == 1 | reported_fy_6m == 1) %>% 
  rename(pc8_result = pc8)



## Prepare result extract for KPI 3.3 analysis ----
### Total number of individuals with an unsatisfacotry result
test_fail_upis <- cervical_results_kpi3 %>%
  filter(reported_fy == 1 & reported_fail_unsat == 1) %>% 
  select(upi_number)

cervical_kpi3_3_visits <- cervical_results_kpi3 %>% 
  # Any result reported within FY + 6 months (to include repeat screen for end of year failed tests)
  filter(reported_fy_6m == 1) %>% 
  # Anyone with a failed test within the FY (including additional records before or after failed test result)
  filter(upi_number %in% pull(test_fail_upis)) %>% 
  mutate(first_reported_fy = reported_fy_6m) %>% 
  # sort by result number to get data in chronological order
  arrange(result_number) %>% 
  group_by(upi_number) %>% 
  # Add new column with numbered visit for each individual
  mutate(visit = 1:n()) %>% 
  ungroup()

### Check how many individuals have a failed test within FY
check_kpi3_visits <- cervical_kpi3_3_visits %>% filter(reported_fail_unsat == 1) 
length(check_kpi3_visits$upi_number)
length(unique(check_kpi3_visits$upi_number))


cervical_results_kpi3_3 <- cervical_kpi3_3_visits %>% 
  arrange(result_number) %>% 
  group_by(upi_number) %>% 
  mutate(first_fail = case_when(reported_fail_unsat == 1 ~ visit),
         first_fail = case_when(is.na(first_fail) ~ max(first_fail, na.rm = TRUE),
                                .default = first_fail)) %>% 
  filter(visit >= first_fail) %>% 
  filter(visit == max(visit)) %>% 
  ungroup() %>% 
  mutate(repeat_scr = ifelse(first_fail < visit, 1, 0))


##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Join extracts ----
##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

## KPI 1.1 Extract, join demographics data with exclusions and results data ----
## Then join with Deprivation data
kpi1_1_extract <- left_join(cervical_demographics_dob,
                                 latest_exclusion, 
                                 by = "upi_number") %>%
  left_join(cervical_results_kpi1_1, by = "upi_number") %>% 
  mutate(pc8_screen = coalesce(pc8_screen, pc8)) %>% 
  left_join(simd, by = c("pc8_screen" = "pc8")) %>% 
  # Exclusion
  mutate(exclusion_flag = ifelse(type_of_exclusion %in% flagged_exclusions, 1, 0)) %>% 
  # Eligible
  cervical_eli_flag() %>%
  # Screened within 3.5 and 5.5 years at time of end of FY
  mutate(scr5_5 = case_when(dateof_examination <= dmy(fy_end_date)
                            & dateof_examination >= dmy(fy_start_date_5_5) ~ 1, 
                            .default = 0),
         scr3_5 = case_when(age <=49 & # Under 50s only
                              dateof_examination <= dmy(fy_end_date) &
                              dateof_examination >= dmy(fy_start_date_3_5) ~ 1, 
                            .default = 0)) %>% 
  mutate(financial_year = fin_yr, .before = 1) %>% 
  mutate(hb2019name = stringr::str_replace(hb2019name, "NHS ", ""),
         simd2020v2_sc_quintile = factor(simd2020v2_sc_quintile,
                                         levels = c("1", "2", "3", "4", "5"))) %>%
  # Drop UPI column as not required any more
  select(!upi_number) %>% 
  select(financial_year,
         "NHS Board of Residence" = hb2019name, 
         "Deprivation Category" = simd2020v2_sc_quintile, 
         dob, age, eligible, scr5_5, scr3_5, exclusion_flag, everything())


## KPI 1.2 Extract, join demographics data with prompt and results data ----
## Then join with Deprivation data
kpi1_2_extract <- right_join(cervical_demographics_dob,
                                  cervical_prompt_data,
                                  by = "upi_number") %>% 
  mutate(pc8_invite = coalesce(pc8_invite, pc8)) %>% 
  left_join(cervical_results_kpi1_2, by = "upi_number") %>% 
  left_join(latest_exclusion, by = "upi_number") %>% 
  left_join(simd, by = c("pc8_invite" = "pc8")) %>% 
  # Invited within FY
  mutate(invite_fy_flag = case_when(age >= 25 & # Remove any <25 from invite_fy_flag as there are very few
                                      (!is.na(I) | !is.na(R1) | !is.na(R2) | !is.na(R3)) ~ 1,
                                    is.na(I) & is.na(R1) & is.na(R2) & is.na(R3) ~ 0,
                                    .default = 0),
         exclusion_flag = ifelse(type_of_exclusion %in% flagged_exclusions, 1, 0)) %>% 
  # Screened within 6 month of a prompt (Invite or Reminder)
  mutate(uptake_flag = case_when(
           between(interval(I, first_doe) %/% days(1), 0, 183) |
             between(interval(I, last_doe) %/% days(1), 0, 183) |
             between(interval(R1, first_doe) %/% days(1), 0, 183) |
             between(interval(R1, last_doe) %/% days(1), 0, 183) |
             between(interval(R2, first_doe) %/% days(1), 0, 183) |
             between(interval(R2, last_doe) %/% days(1), 0, 183) |
             between(interval(R3, first_doe) %/% days(1), 0, 183) |
             between(interval(R3, last_doe) %/% days(1), 0, 183)
           ~ 1,
           .default = 0)) %>% 
  mutate(financial_year = fin_yr, .before = 1) %>% 
  mutate(hb2019name = stringr::str_replace(hb2019name, "NHS ", ""),
         simd2020v2_sc_quintile = factor(simd2020v2_sc_quintile,
                                         levels = c("1", "2", "3", "4", "5"))) %>%
  # Drop UPI column as not required any more
  select(!upi_number) %>% 
  select(financial_year,
         "NHS Board of Residence" = hb2019name, 
         "Deprivation Category" = simd2020v2_sc_quintile, 
         dob, age, invite_fy_flag, uptake_flag, exclusion_flag, everything())


## KPI 2 Extract (use for KPI 3 and KPI 4.1-4.2), join results data with demographics ----
## Then join with Deprivation data
kpi2_extract <- left_join(cervical_results_kpi2,
                          cervical_demographics_dob,
                          by = "upi_number") %>%
  mutate(pc8_result = coalesce(pc8_result, pc8)) %>% 
  left_join(simd, by = c("pc8_result" = "pc8")) %>% 
  mutate(financial_year = fin_yr, .before = 1) %>% 
  mutate(hb2019name = stringr::str_replace(hb2019name, "NHS ", ""),
         simd2020v2_sc_quintile = factor(simd2020v2_sc_quintile,
                                         levels = c("1", "2", "3", "4", "5"))) %>% 
  # Drop UPI column as not required any more
  select(!upi_number) %>% 
  select(financial_year,
         "NHS Board of Residence" = hb2019name, 
         "Deprivation Category" = simd2020v2_sc_quintile,
         dob, age, reported_fy, reported_14, arrived_fy, rejected, everything())


## KPI 3 Extract (KPI 3.1 and KPI 3.2), join results data with demographics ----
## Then join with Deprivation data
kpi3_extract <- left_join(cervical_results_kpi3,
                          cervical_demographics_dob,
                          by = "upi_number") %>% 
  mutate(pc8_result = coalesce(pc8_result, pc8)) %>% 
  left_join(simd, by = c("pc8_result" = "pc8")) %>% 
  mutate(financial_year = fin_yr, .before = 1) %>% 
  mutate(hb2019name = stringr::str_replace(hb2019name, "NHS ", ""),
         simd2020v2_sc_quintile = factor(simd2020v2_sc_quintile,
                                         levels = c("1", "2", "3", "4", "5"))) %>% 
  # Drop UPI column as not required any more
  select(!upi_number) %>% 
  select(financial_year,
         "NHS Board of Residence" = hb2019name, 
         "Deprivation Category" = simd2020v2_sc_quintile,
         dob, age, reported_cyt_fy, reported_cyt_unsat, reported_hpv_fy, reported_hpv_fail, everything())

## KPI 3.3 Extract, join results data with demographics ----
## Then join with Deprivation data
kpi3_3_extract <- left_join(cervical_results_kpi3_3, 
                                 cervical_demographics_dob,
                                 by = "upi_number") %>% 
  mutate(pc8_result = coalesce(pc8_result, pc8)) %>% 
  left_join(simd, by = c("pc8_result" = "pc8")) %>% 
  mutate(financial_year = fin_yr, .before = 1) %>% 
  mutate(hb2019name = stringr::str_replace(hb2019name, "NHS ", ""),
         simd2020v2_sc_quintile = factor(simd2020v2_sc_quintile,
                                         levels = c("1", "2", "3", "4", "5"))) %>% 
  # Drop UPI column as not required any more
  select(!upi_number) %>% 
  select(financial_year,
         "NHS Board of Residence" = hb2019name, 
         "Deprivation Category" = simd2020v2_sc_quintile,
         dob, age, first_reported_fy, repeat_scr, everything())


# Save extracts ----
saveRDS(kpi1_1_extract, file = paste0(fy_folder_path, 
                                      "temp/kpi1_1_extract_", year_label, ".rds"))

saveRDS(kpi1_2_extract, file = paste0(fy_folder_path, 
                                      "temp/kpi1_2_extract_", year_label, ".rds"))

saveRDS(kpi2_extract, file = paste0(fy_folder_path, 
                                    "temp/kpi2_extract_", year_label, ".rds"))

saveRDS(kpi3_extract, file = paste0(fy_folder_path, 
                                    "temp/kpi3_extract_", year_label, ".rds"))

saveRDS(kpi3_3_extract, file = paste0(fy_folder_path, 
                                      "temp/kpi3_3_extract_", year_label, ".rds"))


# End of Script ----
