##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## R Cervical Screening - KPIs for Publication
## 1_housekeeping.R
## 
## Adult Screening Team
## Created: 23/08/2024
## Last update: 09/07/2025
## 
## Written/run on RStudio Posit Workbench
## Platform: x86_64-pc-linux-gnu
## R version 4.4.2 (Pile of Leaves) - 2024-10-31
##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Packages required to source ----
library(lubridate)
library(openxlsx)


# Variables ----
extract_dl_date <- "20250205"
yr_start <- 2023
yr_end <- 2024

first_report_fin_yr <- "2016/17" # for 2022/23 and 2023/24 all historic years will be reported
##!! From 2025/26, this variable can be deleted and updated within script to fin_year_5[1]
##!! Table header for 3.2 and title for Figure 3.2 have hard coded year 2020/21, this should be updated to fin_year_5[1] next year

## Start financial year when 5.5 year-look back is applies across all ages (25-64)
all_scr5_5_start <- 2025 # Reporting period: 2025/26
## Only update if this is changes by the Cervical Screening Programme Board 


## Automated variables
fy_start_date_5_5 <- paste0("01-10-", yr_end-6)
fy_start_date_3_5 <- paste0("01-10-", yr_end-4)
fy_start_date <- paste0("01-04-", yr_start)
fy_end_date <- paste0("31-03-", yr_end)
fy_end_date_plus_6m <- paste0("30-09-", yr_end)
fin_yr <- paste0(yr_start, "/", substr(yr_end, 3, 4))
year_label <- paste0(substr(yr_start, 3, 4), substr(yr_end, 3, 4))
fy_folder <- paste0(yr_start, yr_end)
hist_folder <- paste0(yr_start-1, yr_end-1)

scr_period_1yr <- paste("Reporting period: 1st April", yr_start, "- 31st March", yr_end)

scr_period_5yr <- ifelse(yr_end>all_scr5_5_start,
  paste0("The ", fin_yr, " 5.5-year look-back period: 1st October ", 
         year(dmy(fy_start_date_5_5)), "- 31st March ", yr_end),
  paste0("The ", fin_yr, " 3.5-year look-back period: 1st October ", 
         year(dmy(fy_start_date_3_5)), " - 31st March ", yr_end, 
         " and 5.5-year look-back period: 1st October ", 
         year(dmy(fy_start_date_5_5)), " - 31st March ", yr_end))

fin_year_5 <- c(paste0(yr_end-5, "/", as.double(substr(yr_end, 3, 4))-4),
                paste0(yr_end-4, "/", as.double(substr(yr_end, 3, 4))-3),
                paste0(yr_end-3, "/", as.double(substr(yr_end, 3, 4))-2),
                paste0(yr_end-2, "/", as.double(substr(yr_end, 3, 4))-1),
                fin_yr)


# Import dataset file paths ----
## Set paths
fy_folder_path <- paste0("/PHI_conf/CancerGroup1/Topics/CervicalScreening/AnnualReports/",
                         fy_folder, "/Publication/")
fy_folder_path <- ifelse(yr_start < 2022,
                         paste0("/PHI_conf/CancerGroup1/Topics/CervicalScreening/AnnualReports/",
                                "20162017_20212022", "/Publication/"),
                         fy_folder_path)

hist_folder_path <- paste0("/PHI_conf/CancerGroup1/Topics/CervicalScreening/AnnualReports/",
                           hist_folder, "/Publication/historic/")
hist_folder_path <- ifelse(yr_start < 2023,
                           paste0("/PHI_conf/CancerGroup1/Topics/CervicalScreening/AnnualReports/",
                                  "20162017_20212022", "/Publication/historic/"),
                           hist_folder_path)


## Lookups
simd20_path <- paste0("/conf/linkage/output/lookups/Unicode/Deprivation/",
                      "postcode_2025_1_simd2020v2.rds")

## Saving file paths ----
save_to_kpi1_xlsx <- paste0(fy_folder_path,
                            "output/KPI_1_Coverage_Uptake_", year_label, ".xlsx")
save_to_kpi2_xlsx <- paste0(fy_folder_path,
                            "output/KPI_2_Laboratory_Service_", year_label, ".xlsx")
save_to_kpi3_xlsx <- paste0(fy_folder_path,
                            "output/KPI_3_Screening_Process_", year_label, ".xlsx")


# HB factor levels ----
hbres_levels <- c("Ayrshire & Arran", "Borders", "Dumfries & Galloway",
                  "Fife", "Forth Valley", "Grampian", "Greater Glasgow & Clyde",
                  "Highland", "Lanarkshire", "Lothian", "Orkney", "Shetland",
                  "Tayside", "Western Isles", "Scotland")
hb_levels <- c("S08000015", "S08000016", "S08000017", "S08000029", "S08000019", 
               "S08000020", "S08000031", "S08000022", "S08000032", "S08000024",
               "S08000025", "S08000026", "S08000030", "S08000028", "S92000003")


# KPI Titles for Excel reports ----
kpi1_title <- "KPI 1: Coverage and Uptake"
kpi2_title <- "KPI 2: Laboratory Service"
kpi3_title <- "KPI 3: Screening Process"
kpi4_title <- "KPI 4: Effectiveness of Screening Programme"
kpi5_title <- "KPI 5: Colposcopy Service"
kpi6_title <- "KPI 6: Colposcopy Treatment"


# KPI Notes for Excel reports ----
kpi_notes <- list(
  kpi1_notes = list(
    content = tibble::tibble(
      content = c("Key Performance Indicator (KPI) Report",
            paste("Reporting period:", scr_period_1yr),
            "",
            "  1. Coverage and Uptake",
            paste("      KPI 1.1  Coverage of eligible women who have attended", 
                  "cervical screening within the specified period"),
            paste("      KPI 1.1a Coverage of eligible women who have attended",
                  "screening within the specified period by deprivation (SIMD)"),
            paste("      KPI 1.1b Coverage of eligible women who have attended", 
                  "screening within the specified period by age groups:", 
                  "25-29, 30-34, 35-39, 40-44, 45-49, 50-54, 55-59, 60-64"),
            paste("      KPI 1.1c Coverage of eligible women who have attended", 
                  "screening within the specified period by deprivation (SIMD) and age group"),
            paste("      KPI 1.2  Annual uptake among women who are sent an offer", 
                  "of screening in the latest financial year"),
            paste("      KPI 1.2a Annual uptake among women who are sent an offer", 
                  "of screening in the latest financial year by deprivation (SIMD)"),
            paste("      KPI 1.2b Annual uptake among women who are sent an offer", 
                  "of screening in the latest financial year by age groups:", 
                  "25-29, 30-34, 35-39, 40-44, 45-49, 50-54, 55-59, 60-64"),
            paste("      KPI 1.2c Annual uptake among women who are sent an offer", 
                  "of screening in the latest financial year by deprivation (SIMD) and age group")
    )),
    def = tibble::tibble(
      def = c("Definitions:",
            paste("Coverage is defined as the percentage",
                  "of women in a population eligible for",
                  "screening at a given point in time who",
                  "were screened adequately within a specific period."),
            paste("Uptake is now defined as the proportion",
                  "of those invited for screening who attended",
                  "for a screen within six months (183 days) of",
                  "their invitation date")
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
  ),
  kpi2_notes = list(
    content = tibble::tibble(
      content = c("Key Performance Indicator (KPI) Report",
            paste("Reporting period:", scr_period_1yr),
            "",
            "  2. Laboratory Service",
            paste("      KPI 2.1  Annual percentage of samples reported within", 
                  "two weeks (14 calendar days) from the date of sample being taken"),
            paste("      KPI 2.2  Percentage of tests rejected by the laboratory", 
                  "prior to processing")
      )),
    def = tibble::tibble(
      def = c("Definitions:",
              "Laboratory rejected tests can be for the following reasons:",
              "      Wrong Specimen Type",
              "      Specimen Leaking",
              "      Insufficient Specimen Volume",
              "      Specimen Contaminated",
              "      Out of Date Vial",
              "      Patient Outwith Screening Age",
              "      Wrongly Labelled Vial",
              "      Communication from External Site",
              "      Specimen Not Eligible for Testing",
              "      Other"
      )),
    notes = tibble::tibble(
      notes = c("Notes:",
            paste("Whilst this workbook uses the term ‘women’,",
                  "not only those who identify as women need",
                  "access to cervical screening. Transgender men,",
                  "non-binary and intersex people with a cervix are eligible."),
            paste("On 30 March 2020, Human papillomavirus (HPV) testing replaced", 
                  "cervical cytology as the primary (first) cervical screening test.", 
                  "Cytology based tests are still used but only when HPV is found.")
      )),
    source = tibble::tibble(
      source = c("Source:",
            "Based on the Scottish Cervical Call Recall System (SCCRS).",
            paste0("Data download date: ", substr(extract_dl_date, 7, 8), 
                   "-", substr(extract_dl_date, 5, 6), "-",
                   substr(extract_dl_date, 1, 4))
    ))
  ),
  kpi3_notes = list(
    content = tibble::tibble(
      content = c("Key Performance Indicator (KPI) Report",
                  paste("Reporting period:", scr_period_1yr),
                  "",
                  "  3. Screening Process",
                  paste("      KPI 3.1  Percentage of",
                        "unsatisfactory cytology tests"),
                  paste("      KPI 3.2  Percentage of tests",
                        "processed which were reported as HPV fail"),
                  paste("      KPI 3.3  Percentage of women returning",
                        "for a repeat screen following cytology", 
                        "unsatisfactory or HPV fail")
      )),
    def = tibble::tibble(
      def = c("Definitions:",
              paste("An unsatisfactory cytology test is defined as a test that", 
                    "is not of sufficient quality to enable results to be obtained.")
      )),
    notes = tibble::tibble(
      notes = c("Notes:",
                paste("Whilst this workbook uses the term ‘women’,",
                      "not only those who identify as women need",
                      "access to cervical screening. Transgender men,",
                      "non-binary and intersex people with a cervix are eligible."),
                paste("On 30 March 2020, Human papillomavirus (HPV) testing replaced", 
                      "cervical cytology as the primary (first) cervical screening test.", 
                      "Cytology based tests are still used but only when HPV is found.")
      )),
    source = tibble::tibble(
      source = c("Source:",
                 "Based on the Scottish Cervical Call Recall System (SCCRS).",
                 paste0("Data download date: ", substr(extract_dl_date, 7, 8), 
                        "-", substr(extract_dl_date, 5, 6), "-",
                        substr(extract_dl_date, 1, 4))
    ))
  ),
  kpi4_notes = list(
    content = tibble::tibble(
      x = c("Key Performance Indicator (KPI) Report",
            paste("Reporting period:", scr_period_1yr),
            "",
            "  4. Effectiveness of Screening Programme",
            paste("      KPI 4.1  National annual HPV",
                  "positive rate"),
            paste("      KPI 4.2  Annual colposcopy",
                  "referral rate"),
            paste("      KPI 4.3a Positive predictive value",
                  "of screening result for detection of",
                  "invasive cnancer at colposcopy"),
            paste("      KPI 4.3b Positive predictive value",
                  "of screening result for detection of",
                  "CIN2+ (as the most serious diagnosis)",
                  "at colposcopy"),
            paste("      KPI 4.4a Proportion of cervical",
                  "cancers that are screen detected"),
            paste("      KPI 4.4b Propotion of CIN2+ that",
                  "are screen detected"),
            paste("      KPI 4.5  Proportion of cervical",
                  "cancers that are intercal cancers"),
            paste("      KPI 4.6  Completion of the annual",
                  "National Invasive Cancer Audit (NICA)")
    )),
    def = tibble::tibble(
      x = c()
    ),
    notes = tibble::tibble(
      x = c(),
    ),
    source = tibble::tibble(
      x = c()
  )),
  kpi5_notes = list(
    content = tibble::tibble(
      x = c("Key Performance Indicator (KPI) Report",
            paste("Reporting period:", scr_period_1yr),
            "",
            "  5. Colposcopy Service",
            paste("      KPI 5.1  Percentage of individuals",
                  "on an Urgent, Suspicion of Cancer (USOC)",
                  "referral pathway that were seen by",
                  "colposcopy within 2 weeks of the", 
                  "referral date on SCCRS"),
            paste("      KPI 5.2  Percentage of individuals",
                  "on a high-grade referral pathway that",
                  "were seen by colposcopy within 4 weeks",
                  "of the referral date on SCCRS"),
            paste("      KPI 5.3  Percentage of individuals",
                  "on a low-grade referral pathway that",
                  "were seen by colposcopy within 8 weeks",
                  "of the referral date on SCCRS"),
            paste("      KPI 5.4  Percentage of individuals",
                  "referred to colposcopy that attended",
                  "an appointment"),
            paste("      KPI 5.5  Percentage of individuals",
                  "who receive treatment within 4 weeks",
                  "of diagnostic biopsy")
    )),
    def = tibble::tibble(
      x = c()
    ),
    notes = tibble::tibble(
      x = c(),
    ),
    source = tibble::tibble(
      x = c()
  )),
  kpi6_notes = list(
    content = tibble::tibble(
      x = c("Key Performance Indicator (KPI) Report",
            paste("Reporting period:", scr_period_1yr),
            "",
            "  6. Colposcopy Treatment",
            paste("      KPI 6.1  Percentage of individuals",
                  "with biopsy proven abnormal",
                  "histopathology post treatment"),
            paste("      KPI 6.2  Treatment outcomes:",
                  "Percentage of individuals who have",
                  "undergone treatment with biopsy proven",
                  "evidence of CIN")
    )),
    def = tibble::tibble(
      x = c()
    ),
    notes = tibble::tibble(
      x = c(),
    ),
    source = tibble::tibble(
      x = c()
  ))
)


# KPI 1 Titles, Subtitles and Figure legends ----
kpi1_list <- list(kpi1_title = "KPI 1 - Coverage and Uptake", 
                  kpi1_1_subt = paste("KPI 1.1: Coverage of eligible women who have attended", 
                                      "cervical screening within the specified period"),
                  fig1_1_legend = paste("Figure 1.1: Coverage (%) of eligible women with a previous", 
                                        "screening encounter within the previous 3.5 year or 5.5 years", 
                                        "by NHS Health Board of Residence,", fin_year_5[1], "to", fin_yr),
                  table1_1_header = paste("Table 1.1: Coverage (%) of eligible women", 
                                          "by NHS Health Board of Residence,", first_report_fin_yr, "to", fin_yr),
                  kpi1_1a_subt = paste("KPI 1.1a: Coverage of eligible women who have attended", 
                                       "screening within the specified period by deprivation (SIMD)"),
                  fig1_1a_legend = paste("Figure 1.1a: Scotland Coverage (%) of eligible women with a", 
                                         "previous screening encounter within the previous 3.5 or 5.5 years", 
                                         "by deprivation (SIMD),", fin_year_5[1], "to", fin_yr),
                  table1_1a_hb_header = paste("Table 1.1a1: NHS Health Board of Residence and Scotland", 
                                              "Coverage (%) by deprivation (SIMD), financial year", fin_yr),
                  table1_1a_sc_header = paste("Table 1.1a2: Scotland Coverage (%) by deprivation (SIMD),", 
                                              "financial years,", first_report_fin_yr, "to", fin_yr),
                  
                  kpi1_1b_subt = paste("KPI 1.1b: Coverage of eligible women who have attended",
                                       "screening within the specified period by age groups:", 
                                       "25-29, 30-34, 35-39, 40-44, 45-49, 50-54, 55-59, 60-64"),
                  fig1_1b_legend = paste("Figure 1.1b: Scotland Coverage (%) of eligible women with a", 
                                         "previous screening encounter within the previous 3.5 or 5.5 years", 
                                         "by age group,", fin_year_5[1], "to", fin_yr),
                  table1_1b_hb_header = paste("Table 1.1b1: NHS Health Board of Residence and Scotland", 
                                              "Coverage (%) by age group, financial year", fin_yr),
                  table1_1b_sc_header = paste("Table 1.1b2: Scotland Coverage (%) by age group,", 
                                              "financial years", first_report_fin_yr, "to", fin_yr),
                  
                  kpi1_1c_subt = paste("KPI 1.1c: Coverage of eligible women who have attended", 
                                       "screening within the specified period by deprivation (SIMD) and age group"),
                  fig1_1c_legend_1 = paste("Figure 1.1c1: Scotland Coverage (%) by age group and", 
                                           "deprivation (SIMD), financial year", fin_yr),
                  fig1_1c_legend_2 = paste("Figure 1.1c2: Scotland Coverage (%) by deprivation (SIMD)", 
                                           "and age group, financial year", fin_yr),
                  table1_1c_header = paste("Table 1.1c: Scotland Coverage (%) by deprivation (SIMD)", 
                                           "and age group, financial year", fin_yr),
                  
                  kpi1_2_subt = paste("KPI 1.2: Annual uptake among women who are sent an offer", 
                                      "of screening in the latest financial year"),
                  fig1_2_legend = paste("Figure 1.2: Uptake (%) of women who are sent an offer", 
                                        "of screening within financial years", 
                                        fin_year_5[1], "to", fin_yr, 
                                        "by NHS Health Board of Residence"),
                  table1_2_header = paste("Table 1.2: Uptake (%) by NHS Health Board of Residence,", 
                                          "financial years", first_report_fin_yr, "to", fin_yr),
                  
                  kpi1_2a_subt = paste("KPI 1.2a: Annual uptake among women who are sent an offer", 
                                       "of screening in the latest financial year by deprivation (SIMD)"),
                  fig1_2a_legend = paste("Figure 1.2a: Scotland Uptake (%) of women who are sent an offer", 
                                         "of screening within the financial year, by deprivation (SIMD),",
                                         fin_year_5[1], "to", fin_yr),
                  table1_2a_hb_header = paste("Table 1.2a1: NHS Health Board of Residence and Scotland", 
                                              "Uptake (%) by deprivation (SIMD), financial year", fin_yr),
                  table1_2a_sc_header = paste("Table 1.2a2: Scotland Uptake (%) by deprivation (SIMD),", 
                                              "financial years", first_report_fin_yr, "to", fin_yr),
                  
                  kpi1_2b_subt = paste("KPI 1.2b: Annual uptake among women who are sent an offer", 
                                       "of screening in the latest financial year by age groups:", 
                                       "25-29, 30-34, 35-39, 40-44, 45-49, 50-54, 55-59, 60-64"),
                  fig1_2b_legend = paste("Figure 1.2b: Scotland Uptake (%) of women who are sent an offer", 
                                         "of screening within the financial year, by age group,", 
                                         fin_year_5[1], "to", fin_yr),
                  table1_2b_hb_header = paste("Table 1.2b1: NHS Health Board of Residence and Scotland", 
                                              "Uptake (%) by age group, financial year", fin_yr),
                  table1_2b_sc_header = paste("Table 1.2b2: Scotland Uptake (%) by age group,", 
                                              "financial years", first_report_fin_yr, "to", fin_yr),
                  
                  kpi1_2c_subt = paste("KPI 1.2c: Annual uptake among women who are sent an offer", 
                                       "of screening in the latest financial year by deprivation (SIMD) and age group"),
                  fig1_2c_legend_1 = paste("Figure 1.2c1: Scotland Uptake (%) of women who are sent an offer", 
                                           "of screening within the financial year by age group and deprivation (SIMD),", 
                                           "financial year", fin_yr),
                  fig1_2c_legend_2 = paste("Figure 1.2c2: Scotland Uptake (%) of women who are sent and offer", 
                                           "of screening within the financial year by deprivation (SIMD) and age group,", 
                                           "financial year", fin_yr),
                  table1_2c_header = paste("Table 1.2c: Scotland Uptake (%) by deprivation (SIMD) and age group,", 
                                           "financial year", fin_yr))


# KPI 2 Titles, Subtitles, Figure legends etc ----
kpi2_list <- list(
  kpi2_title = "KPI 2 - Laboratory Service",
  kpi2_1_subt = paste("KPI 2.1: Annual percentage of samples reported within", 
                      "two weeks (14 calendar days) from the date of sample being taken"),
  fig2_1_legend = paste("Figure 2.1: Percentage of sample results reported within", 
                        "two weeks (14 calendar days) by NHS Health Board of Residence,", 
                        fin_year_5[1], "to", fin_yr),
  table2_1_header = paste("Table 2.1: Percentage of sample results reported within", 
                          "two weeks (14 calendar days) by NHS Health Board of Residence,",
                          first_report_fin_yr, "to", fin_yr),
  
  kpi2_2_subt = paste("KPI 2.2: Percentage of tests rejected by the laboratory",
                      "prior to processing"),
  fig2_2_legend = paste("Figure 2.2: Percentage of tests rejected by the laboratory", 
                        "prior to processing, by NHS Health Board of Residence,", 
                        fin_year_5[1], "to", fin_yr),
  table2_2_header = paste("Table 2.2: Percentage of tests rejected by the laboratory", 
                          "prior to processing, by NHS Health Board of Residence,", 
                          first_report_fin_yr, "to", fin_yr),
  footnote = "\u00b9 Human papillomavirus (HPV) testing introduced as primary test."
)


# KPI 3 Titles, Subtitles, Figure legends etc ----
kpi3_list <- list(
  kpi3_title = "KPI 3 - Screening Process",
  kpi3_1_subt = "KPI 3.1: Percentage of unsatisfactory cytology tests",
  fig3_1_legend = paste("Figure 3.1: Percentage of cytology tests reported as Unsatisfactory,",
                        "by NHS Board of Residence,", 
                        head(fin_year_5, n=1), "to", tail(fin_year_5, n=1)),
  table3_1_samples_header = paste("Table 3.1a: Total number of cytology and virology",
                                  "tests processed in Scotland by financial year"),
  table3_1_header = paste("Table 3.1b: Unsatisfactory cytology tests (%) by",
                          "NHS Health Board of Residence,", 
                          first_report_fin_yr, "to", fin_yr),
  kpi3_2_subt = "KPI 3.2: Percentage of virology tests processed which were reported as HPV fail",
  fig3_2_legend = paste("Figure 3.2: Percentage of samples reported as HPV Fail,",
                        "by NHS Board of Residence,", 
                        "2020/21 to", tail(fin_year_5, n=1)),
  table3_2_samples_header = paste("Table 3.2a: Total number of cytology and virology",
                                  "tests processed in Scotland by financial year"),
  table3_2_header = paste("Table 3.2b: Failed HPV tests (%) by NHS Health Board of Residence,", 
                          "2020/21 to", fin_yr),
  kpi3_3_subt = paste("KPI 3.3: Percentage of women returning for a repeat screen",
                      "following cytology unsatisfactory or HPV fail"),
  fig3_3_legend = paste("Figure 3.3: Percentage of women returning for a repeat screen",
                        "following a virology and cytology tests reported as unsatisfactory or HPV fail",
                        "by NHS Board of Residence,", head(fin_year_5, n=1), "to", tail(fin_year_5, n=1)),
  table3_3_header = paste("Table 3.3: Repeated test (%) after a Unsatisfactory or Failed test,",
                          "by NHS Health Board of Residence,", 
                          first_report_fin_yr, "to", fin_yr),
  tablea_footnote = paste("\u00b9 Virology samples introduced when Human papillomavirus (HPV)", 
                          "testing introduced as primary test."),
  tableb_footnote = "\u00b9 Human papillomavirus (HPV) testing introduced as primary test."
)


# KPI 4-6 Title lists to be added ----
# KPI 4 Titles, Subtitles, Figure legends etc
# KPI 5 Titles, Subtitles, Figure legends etc
# KPI 6 Titles, Subtitles, Figure legends etc


# Workbook formatting styles ----
fmt_styles <- list(
  cover_title = createStyle(fontSize = 26, fontColour = "#12436D", halign = "center"),
  cover_title_bg = createStyle(fgFill = "#B3D7F2"),
  b_font = createStyle(textDecoration = "bold"),
  kpi_title_fmt = createStyle(fontSize = 12, fontColour = "#12436D", textDecoration = "bold"),
  halign_right = createStyle(halign = "right"),
  valign_mid = createStyle(valign = "center"),
  thousand_num = createStyle(numFmt = "#,##0", halign = "right"),
  decimal_num = createStyle(numFmt = '[>100]#,###;[=-1]"..";0.0', halign = "right"),
  decimal_two_num = createStyle(numFmt = '[>100]#,###;[=-1]"..";0.00', halign = "right"),
  border_bottom_thick = createStyle(border = "bottom", borderStyle = "medium"),
  font_sz8 = createStyle(fontSize = 8),
  wrap_text = createStyle(wrapText = TRUE),
  halign_centre = createStyle(halign = "center"),
  borders_all = createStyle(border = "bottom-top-left-right")
)


# End of Script ----