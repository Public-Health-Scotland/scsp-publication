##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## R Cervical Screening - KPIs for Publication
## func_cervical_wrt_kpi_data.R
## 
## Adult Screening Team
## Created: 23/08/2024
## Last update: 09/07/2025
## 
## Written/run on RStudio Posit Workbench
## Platform: x86_64-pc-linux-gnu
## R version 4.4.2 (Pile of Leaves) - 2024-10-31
##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


# Function to write to the KPI data sheet for excel reports - One table ----
## wb: workbook with a sheet named Notes
## sheet: string with the name of the Sheet to write the data to
## scr_period: String with the Screening period
## kpi_title: String with the KPI title
## kpi_subt: String with the KPI subtitle
## fig_legend: String with the legend for the figure
## kpi_data: tibble with KPI data
## kpi_fig: Figure presenting KPI data
## styles: Vector with Excel formatting styles

wrt_kpi_1x_table <- function(wb, sheet, scr_period, kpi_subt, fig_legend,
                             kpi_table_header, kpi_data, kpi_fig, styles,
                             footnote) {
  
  # Titles
  writeData(wb, sheet, kpi_subt,
            colNames = FALSE, startCol = 1, startRow = 1)
  writeData(wb, sheet, scr_period,
            colNames = FALSE, startCol = 1, startRow = 2)
  if (grepl("1.1", sheet)) {
    writeData(wb, sheet, "See Notes tab for other look-back periods",
              colNames = FALSE, startCol = 1, startRow = 3)
  }
  
  # Format titles
  addStyle(wb, sheet, styles$kpi_title,
           rows = 1, cols = 1,
           gridExpand = TRUE, stack = TRUE)
  
  # Table title
  writeData(wb, sheet, kpi_table_header,
            startCol = 1, startRow = 5)
  
  # Format Table title
  # addStyle(wb, sheet, styles$phs_magenta_font,
  #          rows = c(2, 5), cols = 1,
  #          gridExpand = TRUE, stack = TRUE)
  
  # Table data
  writeData(wb, sheet, kpi_data,
            startCol = 1, startRow = 6)
  
  # Format table data and header
  addStyle(wb, sheet, styles$halign_right,
           rows = 6, 
           cols = 2:ncol(kpi_data),
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$wrap_text,
           rows = 6, cols = 2:ncol(kpi_data),
           gridExpand = TRUE, stack = TRUE)
  
  if (grepl("KPI 1", sheet)) {
    addStyle(wb, sheet, styles$decimal_num,
             rows = 7:(6+nrow(kpi_data)), cols = 2:ncol(kpi_data),
             gridExpand = TRUE, stack = TRUE)
  } else {
    addStyle(wb, sheet, styles$decimal_two_num,
             rows = 7:(6+nrow(kpi_data)), cols = 2:ncol(kpi_data),
             gridExpand = TRUE, stack = TRUE)
  }
  
  addStyle(wb, sheet, styles$border_bottom_thick,
           rows = c(5, 6, 6+nrow(kpi_data)), cols = 1:ncol(kpi_data),
           gridExpand = TRUE, stack = TRUE)
  
  # Format Table title, table data header and last data row
  addStyle(wb, sheet, styles$b_font,
           rows = c(5, 6, 6+nrow(kpi_data)), 
           cols = 1:ncol(kpi_data),
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$valign_mid,
           rows = c(6, 6+nrow(kpi_data)), 
           cols = 1:ncol(kpi_data),
           gridExpand = TRUE, stack = TRUE)
  
  if (grepl("KPI 2|3", sheet)) {
    writeData(wb, sheet, footnote,
              startCol = 1, 
              startRow = 7+nrow(kpi_data))
    addStyle(wb, sheet, styles$font_sz8,
             rows = 7+nrow(kpi_data), 
             cols = 1)
  }
  
  # Graph
  print(kpi_fig)
  insertPlot(wb, sheet, 
             width = 9, height = 4, 
             startRow = 6, startCol = ncol(kpi_data)+2)
  
  # Graph title
  writeData(wb, sheet, fig_legend,
            startRow = 28, startCol = ncol(kpi_data)+2)
  
  # Format graph title
  addStyle(wb, sheet, styles$b_font,
           rows = 28, cols = ncol(kpi_data)+2,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$wrap_text,
           rows = 28, cols = (ncol(kpi_data)+2):(ncol(kpi_data)+10),
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$halign_centre,
           rows = 28, cols = (ncol(kpi_data)+2):(ncol(kpi_data)+10),
           gridExpand = TRUE, stack = TRUE)
  
  
  # Format worksheet
  setColWidths(wb, sheet, cols = 1, widths = 28)
  setColWidths(wb, sheet, cols = 2:ncol(kpi_data), widths = 12)
  mergeCells(wb, sheet, cols = (ncol(kpi_data)+2):(ncol(kpi_data)+10), rows = 28)
  setRowHeights(wb, sheet, rows = 6, heights = 30)
  setRowHeights(wb, sheet, rows = 6+nrow(kpi_data), heights = 20)
  setRowHeights(wb, sheet, rows = 28, heights = 29)
  
}




# Function to write to the KPI data sheets for excel reports - 2 tables ----
## wb: workbook with a sheet named Notes
## sheet: string with the name of the Sheet to write the data to
## scr_period: String with the Screening period
## kpi_title: String with the KPI title
## kpi_subt: String with the KPI subtitle
## fig_legend: String with the legend for the figure
## kpi_data1: tibble with KPI data, Table 1 (i.e. HB level, 1 yr data)
## kpi_data2: tibble with KPI data, Tabel 2 (i.e. Scotland level, 5 yr data)
## kpi_fig: Figure presenting KPI data
## styles: Vector with Excel formatting styles

wrt_kpi_2x_tables <- function(wb, sheet, scr_period, kpi_subt, fig_legend,
                              kpi_table_1yr_header, kpi_table_5yr_header, 
                              kpi_data1, kpi_data2, kpi_fig, styles, 
                              footnote_a = "", footnote_b = "") {
  
  ## Titles
  writeData(wb, sheet, kpi_subt,
            colNames = FALSE, startCol = 1, startRow = 1)
  writeData(wb, sheet, scr_period,
            colNames = FALSE, startCol = 1, startRow = 2)
  if (grepl("1.1", sheet)) {
    writeData(wb, sheet, "See Notes tab for other look-back periods",
              colNames = FALSE, startCol = 1, startRow = 3)
  }
  
  #  Title format
  addStyle(wb, sheet, styles$kpi_title,
           rows = 1, cols = 1,
           gridExpand = TRUE, stack = TRUE)
  
  
  # Table 1 title
  writeData(wb, sheet, kpi_table_1yr_header,
            startCol = 1, startRow = 5)
  
  # Convert any NAs to -1 to assign .. for no data
  kpi_data1 <- kpi_data1 %>% 
    no_data_reporting_ref()
  
  # Table 1 data
  writeData(wb, sheet, kpi_data1,
            startCol = 1, startRow = 6)
  
  # Format Table 1 data
  if (grepl("KPI 1", sheet)) {
    addStyle(wb, sheet, styles$decimal_num,
             rows = 7:(nrow(kpi_data1)+6), cols = 2:ncol(kpi_data1),
             gridExpand = TRUE, stack = TRUE)
  } else {
    addStyle(wb, sheet, styles$decimal_two_num,
             rows = 7:(nrow(kpi_data1)+6), cols = 2:ncol(kpi_data1),
             gridExpand = TRUE, stack = TRUE)
  }
  
  
  addStyle(wb, sheet, styles$border_bottom_thick,
           rows = c(5, 6, nrow(kpi_data1)+6), 
           cols = 1:ncol(kpi_data1),
           gridExpand = TRUE, stack = TRUE)
  
  if (kpi_data1[1] |> 
      tail(1) |> 
      pull() == "Scotland") {
    
    addStyle(wb, sheet, styles$b_font,
             rows = nrow(kpi_data1)+6, 
             cols = 1:ncol(kpi_data1),
             gridExpand = TRUE, stack = TRUE)
    
  }
  
  if (grepl("3.1|2", sheet)) {
    writeData(wb, sheet, footnote_a,
              startCol = 1, 
              startRow = 7+nrow(kpi_data1))
    addStyle(wb, sheet, styles$font_sz8,
             rows = 7+nrow(kpi_data1), 
             cols = 1)
  }
    
  
  # Write explanation for .. (if any)
  if (kpi_data1 %>%
      select(where(is.double)) %>%
      purrr::map_lgl(~ any(. <0)) %>%
      any == TRUE) {
    
    writeData(wb, sheet, "..  shown where no data is available",
              startRow = 6+nrow(kpi_data1)+4+nrow(kpi_data2)+5, 
              startCol = 1)
  }
  
  # Table 2 title
  writeData(wb, sheet, kpi_table_5yr_header,
            startCol = 1, startRow = 6+nrow(kpi_data1)+3)
  
  # Table 2 data
  writeData(wb, sheet, kpi_data2,
            startCol = 1, startRow = 6+nrow(kpi_data1)+4)
  
  # Format table 2 data
  
  if (grepl("KPI 1", sheet)) {
    addStyle(wb, sheet, styles$decimal_num,
             rows = (6+nrow(kpi_data1)+5):(6+nrow(kpi_data1)+4+nrow(kpi_data2)), 
             cols = 2:ncol(kpi_data2),
             gridExpand = TRUE, stack = TRUE)
  } else {
    addStyle(wb, sheet, styles$decimal_two_num,
             rows = (6+nrow(kpi_data1)+5):(6+nrow(kpi_data1)+4+nrow(kpi_data2)), 
             cols = 2:ncol(kpi_data2),
             gridExpand = TRUE, stack = TRUE)
  }
  
  addStyle(wb, sheet, styles$border_bottom_thick,
           rows = c(6+nrow(kpi_data1)+3, 6+nrow(kpi_data1)+4, 6+nrow(kpi_data1)+4+nrow(kpi_data2)), 
           cols = 1:ncol(kpi_data2),
           gridExpand = TRUE, stack = TRUE)
  
  if (kpi_data2[1] |> 
      tail(1) |> 
      pull() == "Scotland") {
    
    addStyle(wb, sheet, styles$b_font,
             rows = nrow(kpi_data1)+6+nrow(kpi_data2)+4, 
             cols = 1:ncol(kpi_data2),
             gridExpand = TRUE, stack = TRUE)
    
  }
  
  if (grepl("3.1|2", sheet)) {
    writeData(wb, sheet, footnote_b,
              startCol = 1, 
              startRow = 7+nrow(kpi_data1)+4+nrow(kpi_data2))
    addStyle(wb, sheet, styles$font_sz8,
             rows = 7+nrow(kpi_data1)+4+nrow(kpi_data2), 
             cols = 1)
  }
  
  #Format Table 1+2 data header + Titles
  addStyle(wb, sheet, styles$b_font,
           rows = c(5, 6, 
                    5+nrow(kpi_data1)+4, 6+nrow(kpi_data1)+4), 
           cols =  1:max(ncol(kpi_data1), ncol(kpi_data2)),
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$valign_mid,
           rows = c(6, 6+nrow(kpi_data1)+4), 
           cols = 1:max(ncol(kpi_data1), ncol(kpi_data2)),
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$wrap_text,
           rows = c(6, 6+nrow(kpi_data1)+4), 
           cols = 1:max(ncol(kpi_data1), ncol(kpi_data2)),
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$halign_centre,
           rows = c(6:(6+nrow(kpi_data1)), 
                    (6+nrow(kpi_data1)+4):(6+nrow(kpi_data1)+4+nrow(kpi_data2))), 
           cols = 1:max(ncol(kpi_data1), ncol(kpi_data2)),
           gridExpand = TRUE, stack = TRUE)
  # addStyle(wb, sheet, styles$phs_magenta_font,
  #          rows = c(2, 5, 6+nrow(kpi_data1)+3), cols = 1,
  #          gridExpand = TRUE, stack = TRUE)
  
  
  # Graph
  print(kpi_fig)
  insertPlot(wb, sheet, 
             width = 7.5, height = 4, 
             startRow = 6+nrow(kpi_data1)+4, startCol = ncol(kpi_data2)+2)
  # Graph Title
  writeData(wb, sheet, fig_legend,
            startRow = 6+nrow(kpi_data1)+4+22, 
            startCol = ncol(kpi_data2)+2)
  
  # Format Graph Title
  addStyle(wb, sheet, styles$b_font,
           rows = 6+nrow(kpi_data1)+4+22,
           cols = ncol(kpi_data2)+2,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$wrap_text,
           rows = 6+nrow(kpi_data1)+4+22, 
           cols = (ncol(kpi_data2)+2):(ncol(kpi_data2)+2+5),
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$halign_centre,
           rows = 6+nrow(kpi_data1)+4+22, 
           cols = (ncol(kpi_data2)+2):(ncol(kpi_data2)+2+5),
           gridExpand = TRUE, stack = TRUE)
  
  # Format work sheet
  setColWidths(wb, sheet, cols = 1, widths = 28)
  setColWidths(wb, sheet, cols = 2:ncol(kpi_data1), widths = 15)
  mergeCells(wb, sheet, rows = 6+nrow(kpi_data1)+4+22,
             cols = (ncol(kpi_data2)+2):(ncol(kpi_data2)+2+5))
  setRowHeights(wb, sheet, rows = 6, heights = 32) 
  setRowHeights(wb, sheet, rows = 6+nrow(kpi_data1)+4, heights = 32)
  setRowHeights(wb, sheet, rows = 6+nrow(kpi_data1)+4+22, heights = 29)
  
}


# Function to write to the KPI data sheets for excel reports - 1 table, 2 graphs ----
## wb: workbook with a sheet named Notes
## sheet: string with the name of the Sheet to write the data to
## scr_period: String with the Screening period
## kpi_subt: String with the KPI subtitle
## fig_legend1: String with the legend for figure 1
## fig_legend2: String with the legend for the figure 2
## kpi_data: tibble with KPI data, Scotland-level data
## kpi_fig1: Figure 1 presenting KPI data
## kpi_fig2: Figure 2 presenting KPI data
## styles: Vector with Excel formatting styles

wrt_kpi_2x_graphs <- function(wb, sheet, scr_period, kpi_subt, 
                              fig_legend1, fig_legend2, kpi_table_header, 
                              kpi_data, kpi_fig1, kpi_fig2, styles) {
  
  # Titles
  writeData(wb, sheet, kpi_subt,
            colNames = FALSE, startCol = 1, startRow = 1)
  writeData(wb, sheet, scr_period,
            colNames = FALSE, startCol = 1, startRow = 2)
  if (grepl("1.1", sheet)) {
    writeData(wb, sheet, "See Notes tab for other look-back periods",
              colNames = FALSE, startCol = 1, startRow = 3)
  }
  
  # Format title
  addStyle(wb, sheet, styles$kpi_title,
           rows = 1, cols = 1,
           gridExpand = TRUE, stack = TRUE)
  # addStyle(wb, sheet, styles$phs_magenta_font,
  #          rows = c(2, 5), cols = 1,
  #          gridExpand = TRUE, stack = TRUE)
  
  
  # Table title
  writeData(wb, sheet, kpi_table_header,
            startCol = 1, startRow = 5)
  
  # Convert any NAs to -1 to assign .. for no data
  kpi_data <- kpi_data %>% 
    no_data_reporting_ref()
  
  # Table data
  writeData(wb, sheet, kpi_data,
            startCol = 1, startRow = 6)
  
  # Format Table title
  addStyle(wb, sheet, styles$b_font,
           rows = 6, cols = 2:ncol(kpi_data),
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$wrap_text,
           rows = 6, cols = 1:ncol(kpi_data),
           gridExpand = TRUE, stack = TRUE)
  
  # Format Table title and data
  addStyle(wb, sheet, styles$valign_mid,
           rows = c(6, 6+nrow(kpi_data)), cols = 1:ncol(kpi_data),
           gridExpand = TRUE, stack = TRUE)
  
  if (grepl("KPI 1", sheet)) {
    addStyle(wb, sheet, styles$decimal_num,
             rows = 7:(nrow(kpi_data)+6), cols = 2:ncol(kpi_data),
             gridExpand = TRUE, stack = TRUE)
  } else {
    addStyle(wb, sheet, styles$decimal_two_num,
             rows = 7:(nrow(kpi_data)+6), cols = 2:ncol(kpi_data),
             gridExpand = TRUE, stack = TRUE)
  }
  
  addStyle(wb, sheet, styles$border_bottom_thick,
           rows = c(5, 6, nrow(kpi_data)+6), cols = 1:ncol(kpi_data),
           gridExpand = TRUE, stack = TRUE)
  
  addStyle(wb, sheet, styles$halign_centre,
           rows = 6:(6+nrow(kpi_data)), 
           cols = 1:ncol(kpi_data),
           gridExpand = TRUE, stack = TRUE)
  
  # Write explanation for .. (if any)
  if (kpi_data %>%
      select(where(is.double)) %>%
      purrr::map_lgl(~ any(. <0)) %>%
      any == TRUE) {
    
    writeData(wb, sheet, "..  shown where no data is available",
              startRow = 6+nrow(kpi_data)+4+nrow(kpi_data)+5, 
              startCol = 1)
  }
  
  # Graph 1
  print(kpi_fig1)
  insertPlot(wb, sheet, 
             width = 7.5, height = 4, 
             startRow = 6+nrow(kpi_data)+4, startCol = 1)
  
  # Graph 2
  print(kpi_fig2)
  insertPlot(wb, sheet, 
             width = 7.5, height = 4, 
             startRow = 6+nrow(kpi_data)+4, startCol = 8)
  
  # Graph title 1
  writeData(wb, sheet, fig_legend1,
            startRow = 6+nrow(kpi_data)+4+23, 
            startCol = 1)
  
  # Graph titel 2
  writeData(wb, sheet, fig_legend2,
            startRow = 6+nrow(kpi_data)+4+23, 
            startCol = 8)
  
  # Format Graphs 1+2 titles
  addStyle(wb, sheet, styles$b_font,
           rows = c(5:(6+nrow(kpi_data)+4+23)), cols = 1,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$b_font,
           rows = 6+nrow(kpi_data)+4+23, cols = 8,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$wrap_text,
           rows = 6+nrow(kpi_data)+4+23, 
           cols = c(1:6, 8:15),
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$halign_centre,
           rows = 6+nrow(kpi_data)+4+23, 
           cols = c(1:6, 8:15),
           gridExpand = TRUE, stack = TRUE)
  
  
  # Format worksheet
  setColWidths(wb, sheet, cols = 1, widths = 19)
  setColWidths(wb, sheet, cols = 2:ncol(kpi_data), widths = 15)
  mergeCells(wb, sheet, rows = 6+nrow(kpi_data)+4+23,
             cols = 1:6)
  mergeCells(wb, sheet, rows = 6+nrow(kpi_data)+4+23,
             cols = 8:15)
  setRowHeights(wb, sheet, rows = 6, heights = 32)
  setRowHeights(wb, sheet, rows = 6+nrow(kpi_data)+4+23, heights = 40)
  
  
}












