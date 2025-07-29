##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## R Cervical Screening - KPIs for Publication
## func_cervical_wrt_xlsx_workbook_2549_5064.R
## 
## Adult Screening Team
## Created: 23/08/2024
## Last update: 09/07/2025
## 
## Written/run on RStudio Posit Workbench
## Platform: x86_64-pc-linux-gnu
## R version 4.4.2 (Pile of Leaves) - 2024-10-31
##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


# Function to write the KPI 1 excel report, 3.5 and 5.5 split tabs ----
## wb: workbook with a sheet named Notes
## scr_1yr: The 1 year annual screening period
## scr_5yr: The 5 year screening look-back period
## styles: Vector with Excel formatting styles 
## kpi_notes: List notes for the KPI
## kpi_list: KPI specific List of titles, legends etc. 
## kpix_x_data: Tibble with the KPI data for each of the sub-KPIs
## kpix_x_fig: Figure presenting data for each of the sub-KPIs
## wb_file_name: File name of the workbook to be saved as

wrt_xlsx_kpi1_2549_5064 <- function(scr_1yr, scr_3yr, scr_5yr, styles, kpi_notes, kpi1_list,
                                    kpi1_1_data_2549, kpi1_1_fig_2549,
                                    kpi1_1a_data_1yr_2549, kpi1_1a_data_5yr_2549, kpi1_1a_fig_2549, 
                                    kpi1_1_data_5064, kpi1_1_fig_5064,
                                    kpi1_1a_data_1yr_5064, kpi1_1a_data_5yr_5064, kpi1_1a_fig_5064,
                                    wb_file_name) {
  
  wb <- createWorkbook()
  
  options("openxlsx.dateFormat" = "dd/mm/yyyy",
          "openxlsx.gridLines" = FALSE)
  modifyBaseFont(wb, fontSize = 10,
                 fontName = "Arial")
  
  # Add worksheets
  addWorksheet(wb, sheetName = "Cover")
  addWorksheet(wb, sheetName = "Notes")
  addWorksheet(wb, sheetName = "KPI 1.1 - Coverage, 3.5yr")
  addWorksheet(wb, sheetName = "KPI 1.1 - Coverage, 5.5yr")
  addWorksheet(wb, sheetName = "KPI 1.1a - Coverage SIMD, 3.5yr")
  addWorksheet(wb, sheetName = "KPI 1.1a - Coverage SIMD, 5.5yr")
  
  # Write Cover and Notes sheets
  wrt_xlsx_cover(wb, scr_1yr, kpi1_list$kpi1_title, styles)
  wrt_xlsx_notes(wb, kpi_notes, styles)
  
  # Write KPI data
  ## KPI 1.1
  wrt_kpi_5yr(wb, "KPI 1.1 - Coverage, 3.5yr", scr_3yr, 
              kpi1_list$kpi1_1_subt_35, kpi1_list$fig1_1_legend_35,
              kpi1_list$table1_1_header_35, kpi1_1_data_2549, kpi1_1_fig_2549, styles)
  wrt_kpi_5yr(wb, "KPI 1.1 - Coverage, 5.5yr", scr_5yr, 
              kpi1_list$kpi1_1_subt_55, kpi1_list$fig1_1_legend_55,
              kpi1_list$table1_1_header_55, kpi1_1_data_5064, kpi1_1_fig_5064, styles)
  
  ## KPI 1.1a
  wrt_kpi_1yr_5yr(wb, "KPI 1.1a - Coverage SIMD, 3.5yr", scr_3yr, 
                  kpi1_list$kpi1_1a_subt_35, kpi1_list$fig1_1a_legend_35,
                  kpi1_list$table1_1a_hb_header_35, kpi1_list$table1_1a_sc_header_35, 
                  kpi1_1a_data_1yr_2549, kpi1_1a_data_5yr_2549, kpi1_1a_fig_2549, styles)
  wrt_kpi_1yr_5yr(wb, "KPI 1.1a - Coverage SIMD, 5.5yr", scr_5yr, 
                  kpi1_list$kpi1_1a_subt_55, kpi1_list$fig1_1a_legend_55,
                  kpi1_list$table1_1a_hb_header_55, kpi1_list$table1_1a_sc_header_55, 
                  kpi1_1a_data_1yr_5064, kpi1_1a_data_5yr_5064, kpi1_1a_fig_5064, styles)
  
  
  # Save workbook
  saveWorkbook(wb, wb_file_name, overwrite = TRUE)
}



