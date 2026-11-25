################################################################################
## R Cervical Screening Publication - KPI 1-3
## func_cervical_wrt_kpi_data.R
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

#### Function to write to the Cover sheet for excel reports
## wb: workbook with a sheet named Cover
## scr_period: The screening period, i.e. financial year: 1st April - 31st March
## kpi_title: KPI specfic title variable
## styles: Vector with Excel formatting styles 

wrt_xlsx_cover <- function(wb, scr_period, kpi_title, styles) {
  
  cover_title <- tibble::tibble(x = c("Scottish Cervical Screening Programme",
                                      "Key Performance Indicator Report",
                                      scr_period,
                                      "",
                                      kpi_title))
  
  writeData(wb, sheet = "Cover", cover_title,
            colNames = FALSE, startCol = 7, startRow = 5)
  
  addStyle(wb, sheet = "Cover", styles$cover_title,
           rows = 5:9, cols = 7,
           gridExpand = TRUE, stack = TRUE)
  
  addStyle(wb, sheet = "Cover", styles$b_font,
           rows = c(5, 9), cols = 7,
           gridExpand = TRUE, stack = TRUE)
  
  addStyle(wb, sheet = "Cover", styles$cover_title_bg,
           rows = 2:12, cols = 2:12,
           gridExpand = TRUE, stack = TRUE)
  
}




#### Function to write to the Notes sheet for excel reports
## wb: workbook with a sheet named Notes
## kpi_notes: List notes for the KPI
## styles: Vector with Excel formatting styles 

wrt_xlsx_notes <- function(wb, kpi_notes, styles) {
  row_counter <- 1
  
  writeData(wb, sheet = "Notes", kpi_notes$content,
            colNames = FALSE, startCol = 1, 
            startRow = row_counter)
  
  # Format Notes
  addStyle(wb, sheet = "Notes", styles$kpi_title,
           rows = row_counter, cols = 1,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet = "Notes", styles$b_font,
           rows = c(row_counter + 1, row_counter + 3), cols = 1,
           gridExpand = TRUE, stack = TRUE)
  # addStyle(wb, sheet = "Notes", styles$phs_blue_font,
  #          rows = 4, cols = 1,
  #          gridExpand = TRUE, stack = TRUE)

  row_counter <- row_counter + nrow(kpi_notes$content) + 2

  if ("def" %in% names(kpi_notes)){
  writeData(wb, sheet = "Notes", kpi_notes$def,
            colNames = FALSE, startCol = 1, 
            startRow = row_counter)
  # Format Def text
  addStyle(wb, sheet = "Notes", styles$kpi_title,
           rows = row_counter, cols = 1,
           gridExpand = TRUE, stack = TRUE)
  
  row_counter <- row_counter + nrow(kpi_notes$def) + 2
  }
  
  if ("notes" %in% names(kpi_notes)){
  writeData(wb, sheet = "Notes", kpi_notes$notes,
            colNames = FALSE, startCol = 1, 
            startRow = row_counter)
  
    # Format Notes text
  addStyle(wb, sheet = "Notes", styles$kpi_title,
           rows = row_counter, cols = 1,
           gridExpand = TRUE, stack = TRUE)
  
  row_counter <- row_counter + nrow(kpi_notes$notes)
  if ("table1" %in% names(kpi_notes)){
    writeData(wb, sheet = "Notes", kpi_notes$table1,
              colNames = FALSE, startCol = 2, 
              startRow = row_counter)

    writeData(wb, sheet = "Notes", kpi_notes$table2,
              colNames = FALSE, startCol = 7, 
              startRow = row_counter)
    
    # Format look-back period table
    # addStyle(wb, sheet = "Notes", styles$b_font,
    #          rows = nrow(kpi_notes$content)+2+nrow(kpi_notes$def)+1+nrow(kpi_notes$notes)-1, 
    #          cols = c(1,7),
    #          gridExpand = TRUE, stack = TRUE)

    addStyle(wb, sheet = "Notes", styles$b_font,
             rows = row_counter, 
             cols = c(2:5, 7:8),
             gridExpand = TRUE, stack = TRUE)

    addStyle(wb, sheet = "Notes", styles$wrap_text,
             rows = row_counter:(row_counter + nrow(kpi_notes$table1) - 1),
             cols = c(2:5, 7:8),
             gridExpand = TRUE, stack = TRUE)

    addStyle(wb, sheet = "Notes", styles$valign_mid,
             rows = row_counter:(row_counter + nrow(kpi_notes$table1) - 1),
             cols = c(2:5, 7:8),
             gridExpand = TRUE, stack = TRUE)

    addStyle(wb, sheet = "Notes", styles$borders_all,
             rows = row_counter:(row_counter + nrow(kpi_notes$table1) - 1),
             cols = c(2:5, 7:8),
             gridExpand = TRUE, stack = TRUE)

    row_counter <- row_counter + nrow(kpi_notes$table1) + 2
  } else {
    row_counter <- row_counter + 2
  }
  }

  if ("source" %in% names(kpi_notes)){
    writeData(wb, sheet = "Notes", kpi_notes$source,
              colNames = FALSE, startCol = 1, 
              startRow = row_counter)
    
    # Format Content text
    # addStyle(wb, sheet = "Notes", styles$phs_magenta_font,
    #          rows = 5:nrow(kpi_notes$content), cols = 1,
    #          gridExpand = TRUE, stack = TRUE)

    # Format Source text
    addStyle(wb, sheet = "Notes", styles$kpi_title,
             rows = row_counter, cols = 1,
             gridExpand = TRUE, stack = TRUE)
  }
  
  # Set Column width
  setColWidths(wb, sheet= "Notes", cols = c(2, 4, 7), widths = 19)
  setColWidths(wb, sheet= "Notes", cols = c(5,8), widths = 31)
  
  
}




#### Function to write the KPI 1 excel report
## wb: workbook with a sheet named Notes
## scr_1yr: The 1 year annual screening period
## scr_5yr: The 5 year screening look-back period
## styles: Vector with Excel formatting styles 
## kpi_notes: List notes for the KPI
## kpi_list: KPI specific List of titles, legends etc. 
## kpix_x_data: Tibble with the KPI data for each of the sub-KPIs
## kpix_x_fig: Figure presenting data for each of the sub-KPIs
## wb_file_name: File name of the workbook to be saved as

wrt_xlsx_kpi1 <- function(scr_1yr, scr_5yr, styles, kpi_notes, kpi1_list,
                          kpi1_1_data, kpi1_1_scot_fig, kpi1_1_fig, 
                          kpi1_1a_data_1yr, kpi1_1a_data_5yr, kpi1_1a_fig, 
                          kpi1_1b_data_1yr, kpi1_1b_data_5yr, kpi1_1b_fig,
                          kpi1_1c_data, kpi1_1c_fig1, kpi1_1c_fig2,
                          #kpi1_1d_data, kpi1_1d_fig,
                          kpi1_2_data, kpi1_2_scot_fig, kpi1_2_fig, 
                          kpi1_2a_data_1yr, kpi1_2a_data_5yr, kpi1_2a_fig, 
                          kpi1_2b_data_1yr, kpi1_2b_data_5yr, kpi1_2b_fig,
                          kpi1_2c_data, kpi1_2c_fig1, kpi1_2c_fig2,
                          #kpi1_2d_data, kpi1_2d_fig,
                          wb_file_name) {
  
  wb <- createWorkbook()
  
  options("openxlsx.dateFormat" = "dd/mm/yyyy",
          "openxlsx.gridLines" = FALSE)
  modifyBaseFont(wb, fontSize = 10,
                 fontName = "Arial")
  
  # Add worksheets
  addWorksheet(wb, sheetName = "Cover")
  addWorksheet(wb, sheetName = "Notes")
  addWorksheet(wb, sheetName = "KPI 1.1 - Coverage")
  addWorksheet(wb, sheetName = "KPI 1.1a - Coverage by SIMD")
  addWorksheet(wb, sheetName = "KPI 1.1b - Coverage by age")
  addWorksheet(wb, sheetName = "KPI 1.1c - Coverage by SIMD+age")
  #addWorksheet(wb, sheetName = "KPI 1.1d - Coverage by Vacc")
  addWorksheet(wb, sheetName = "KPI 1.2 - Uptake")
  addWorksheet(wb, sheetName = "KPI 1.2a - Uptake by SIMD")
  addWorksheet(wb, sheetName = "KPI 1.2b - Uptake by age")
  addWorksheet(wb, sheetName = "KPI 1.2c - Uptake by SIMD+age")
  #addWorksheet(wb, sheetName = "KPI 1.2d - Uptake by Vacc")
  
  # Write Cover and Notes sheets
  wrt_xlsx_cover(wb, scr_1yr, kpi1_list$kpi1_title, styles)
  wrt_xlsx_notes(wb, kpi_notes, styles)
  
  # Write KPI data
  ## KPI 1.1
  wrt_kpi_2x_graphs(wb, "KPI 1.1 - Coverage", scr_5yr,
                   kpi1_list$kpi1_1_subt, kpi1_list$fig1_1_scot_legend, kpi1_list$fig1_1_legend, 
                   kpi1_list$table1_1_header, kpi1_1_data, kpi1_1_scot_fig, kpi1_1_fig, styles,
                   plot_1_width = 10, plot_2_width = 10)

  ## KPI 1.1a
  wrt_kpi_2x_tables(wb, "KPI 1.1a - Coverage by SIMD", scr_5yr,
                    kpi1_list$kpi1_1a_subt, kpi1_list$fig1_1a_legend,
                    kpi1_list$table1_1a_hb_header, kpi1_list$table1_1a_sc_header,
                    kpi1_1a_data_1yr, kpi1_1a_data_5yr, kpi1_1a_fig, styles)

  ## KPI 1.1b
  wrt_kpi_2x_tables(wb, "KPI 1.1b - Coverage by age", scr_5yr,
                    kpi1_list$kpi1_1b_subt, kpi1_list$fig1_1b_legend,
                    kpi1_list$table1_1b_hb_header, kpi1_list$table1_1b_sc_header,
                    kpi1_1b_data_1yr, kpi1_1b_data_5yr, kpi1_1b_fig, styles)

  ## KPI 1.1c
  wrt_kpi_2x_graphs(wb, "KPI 1.1c - Coverage by SIMD+age", scr_5yr, kpi1_list$kpi1_1c_subt,
                    kpi1_list$fig1_1c_legend_1, kpi1_list$fig1_1c_legend_2,
                    kpi1_list$table1_1c_header, kpi1_1c_data, kpi1_1c_fig1, kpi1_1c_fig2, styles)

  # ## KPI 1.1d
  # wrt_kpi_2x_tables(wb, "KPI 1.1d - Coverage by Vacc", scr_5yr,
  #                   kpi1_list$kpi1_1d_subt, kpi1_list$fig1_1d_legend,
  #                   kpi1_list$table1_1d_header_1, kpi1_list$table1_1d_header_2,
  #                   kpi1_1d_data[1,2][[1]][[1]], kpi1_1d_data[2,2][[1]][[1]],
  #                   kpi1_1d_fig, styles, col_width = 7)
  
  ## KPI 1.2
  wrt_kpi_2x_graphs(wb, "KPI 1.2 - Uptake", scr_1yr,
                    kpi1_list$kpi1_2_subt, kpi1_list$fig1_2_scot_legend, kpi1_list$fig1_2_legend, 
                    kpi1_list$table1_2_header, kpi1_2_data, kpi1_2_scot_fig, kpi1_2_fig, styles,
                    plot_1_width = 10, plot_2_width = 10)
  ## KPI 1.2a
  wrt_kpi_2x_tables(wb, "KPI 1.2a - Uptake by SIMD", scr_1yr,
                    kpi1_list$kpi1_2a_subt, kpi1_list$fig1_2a_legend,
                    kpi1_list$table1_2a_hb_header, kpi1_list$table1_2a_sc_header,
                    kpi1_2a_data_1yr, kpi1_2a_data_5yr, kpi1_2a_fig, styles)

  ## KPI 1.2b
  wrt_kpi_2x_tables(wb, "KPI 1.2b - Uptake by age", scr_1yr,
                    kpi1_list$kpi1_2b_subt, kpi1_list$fig1_2b_legend,
                    kpi1_list$table1_2b_hb_header, kpi1_list$table1_2b_sc_header,
                    kpi1_2b_data_1yr, kpi1_2b_data_5yr, kpi1_2b_fig, styles)

  ## KPI 1.2c
  wrt_kpi_2x_graphs(wb, "KPI 1.2c - Uptake by SIMD+age", scr_1yr, kpi1_list$kpi1_2c_subt,
                    kpi1_list$fig1_2c_legend_1, kpi1_list$fig1_2c_legend_2,
                    kpi1_list$table1_2c_header, kpi1_2c_data, kpi1_2c_fig1, kpi1_2c_fig2, styles)

  # ## KPI 1.2d
  # wrt_kpi_2x_tables(wb, "KPI 1.2d - Uptake by Vacc", scr_5yr,
  #                   kpi1_list$kpi1_2d_subt, kpi1_list$fig1_2d_legend,
  #                   kpi1_list$table1_2d_header_1, kpi1_list$table1_2d_header_2,
  #                   kpi1_2d_data[1,2][[1]][[1]], kpi1_2d_data[2,2][[1]][[1]],
  #                   kpi1_2d_fig,
  #                   styles)
  # Save workbook
  saveWorkbook(wb, wb_file_name, overwrite = TRUE)
}



#### Function to write the KPI 2 excel report
## wb: workbook with a sheet named Notes
## scr_1yr: The 1 year annual screening period
## styles: Vector with Excel formatting styles 
## kpi_notes: List notes for the KPI
## kpi_list: KPI specific List of titles, legends etc. 
## kpix_x_data: Tibble with the KPI data for each of the sub-KPIs
## kpix_x_fig: Figure presenting data for each of the sub-KPIs
## wb_file_name: File name of the workbook to be saved as

wrt_xlsx_kpi2 <- function(wb, scr_1yr, styles, kpi_notes, kpi_list, 
                          kpi2_1_data, kpi2_1_fig, 
                          kpi2_2_data, kpi2_2_fig,
                          wb_file_name) {
  
  wb <- createWorkbook()
  
  options("openxlsx.dateFormat" = "dd/mm/yyyy",
          "openxlsx.gridLines" = FALSE)
  modifyBaseFont(wb, fontSize = 10,
                 fontName = "Arial")
  
  # Add worksheets
  addWorksheet(wb, sheetName = "Cover")
  addWorksheet(wb, sheetName = "Notes")
  addWorksheet(wb, sheetName = "KPI 2.1 - Reporting")
  addWorksheet(wb, sheetName = "KPI 2.2 - Rejected Samples")

  # Write Cover and Notes sheets
  wrt_xlsx_cover(wb, scr_1yr, kpi_list$kpi2_title, styles)
  wrt_xlsx_notes(wb, kpi_notes, styles)
  
  # Write KPI data
  ## KPI 2.1
  wrt_kpi_1x_table(wb, "KPI 2.1 - Reporting", scr_1yr,
                   kpi_list$kpi2_1_subt, kpi2_list$fig2_1_legend,
                   kpi_list$table2_1_header, kpi2_1_data, kpi2_1_fig, styles,
                   footnote = kpi2_list$footnote)
  
  ## KPI 2.2
  wrt_kpi_1x_table(wb, "KPI 2.2 - Rejected Samples", scr_1yr,
                   kpi_list$kpi2_2_subt, kpi2_list$fig2_2_legend,
                   kpi_list$table2_2_header, kpi2_2_data, kpi2_2_fig, styles,
                   footnote = kpi2_list$footnote, decimal_points = 2)

  # Save workbook
  saveWorkbook(wb, wb_file_name, overwrite = TRUE)
}




#### Function to write the KPI 3 excel report
## wb: workbook with a sheet named Notes
## scr_1yr: The 1 year annual screening period
## styles: Vector with Excel formatting styles 
## kpi_notes: List notes for the KPI
## kpi_list: KPI specific List of titles, legends etc. 
## kpix_x_data: Tibble with the KPI data for each of the sub-KPIs
## kpix_x_fig: Figure presenting data for each of the sub-KPIs
## wb_file_name: File name of the workbook to be saved as

wrt_xlsx_kpi3 <- function(wb, scr_1yr, styles, kpi_notes, kpi_list, 
                          kpi3_data_samples, 
                          kpi3_1_data, kpi3_1_fig, 
                          kpi3_2_data, kpi3_2_fig,
                          kpi3_3_data, kpi3_3_fig,
                          wb_file_name) {
  
  wb <- createWorkbook()
  
  options("openxlsx.dateFormat" = "dd/mm/yyyy",
          "openxlsx.gridLines" = FALSE)
  modifyBaseFont(wb, fontSize = 10,
                 fontName = "Arial")
  
  # Add worksheets
  addWorksheet(wb, sheetName = "Cover")
  addWorksheet(wb, sheetName = "Notes")
  addWorksheet(wb, sheetName = "KPI 3.1 - Cytology Tests")
  addWorksheet(wb, sheetName = "KPI 3.2 - HPV Tests")
  addWorksheet(wb, sheetName = "KPI 3.3 - Repeated Tests")
  
  # Write Cover and Notes sheets
  wrt_xlsx_cover(wb, scr_1yr, kpi_list$kpi3_title, styles)
  wrt_xlsx_notes(wb, kpi_notes, styles)
  
  # Write KPI data
  ## KPI 3.1
  wrt_kpi_2x_tables(wb, "KPI 3.1 - Cytology Tests", scr_1yr, 
                    kpi_list$kpi3_1_subt, kpi_list$fig3_1_legend,
                    kpi_list$table3_1_samples_header, kpi_list$table3_1_header,
                    kpi3_data_samples, kpi3_1_data, kpi3_1_fig, styles,
                    footnote_a = kpi_list$tablea_footnote,
                    footnote_b = kpi_list$tableb_footnote)
  
  ## KPI 3.2
  wrt_kpi_2x_tables(wb, "KPI 3.2 - HPV Tests", scr_1yr,
                    kpi_list$kpi3_2_subt, kpi_list$fig3_2_legend,
                    kpi_list$table3_2_samples_header, kpi_list$table3_2_header,
                    kpi3_data_samples, kpi3_2_data, kpi3_2_fig, styles,
                    footnote_a = kpi_list$tablea_footnote,
                    footnote_b = kpi_list$tableb_footnote,
                    decimal_points = 2)
  
  ## KPI 3.3
  wrt_kpi_1x_table(wb, "KPI 3.3 - Repeated Tests", scr_1yr,
                   kpi_list$kpi3_3_subt, kpi_list$fig3_3_legend,
                   kpi_list$table3_3_header, kpi3_3_data, kpi3_3_fig, styles,
                   footnote = kpi_list$tableb_footnote)
  
  # Save workbook
  saveWorkbook(wb, wb_file_name, overwrite = TRUE)
}
