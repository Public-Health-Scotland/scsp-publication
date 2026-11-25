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

wrt_kpi_table <- function(wb, sheet,
                          col_start, row_start,
                          kpi_table_header, table,
                          styles,
                          footnote=NULL,
                          table_col_center_align = 0,
                          decimal_points = 1,
                          ...){
  
  table_length <- nrow(table)
  if (names(table)[1] == "Health Board of Residence") {
    hb_table <- TRUE
  } else {
    hb_table <- FALSE
  }
  
  # Table title
  writeData(wb, sheet, kpi_table_header,
            startCol = col_start, startRow = row_start)
  setRowHeights(wb, sheet, rows = row_start + 1, heights = 24) 
  addStyle(wb, sheet, styles$valign_bottom,
           rows = row_start + 1, 
           cols = col_start,
           gridExpand = TRUE, stack = TRUE)
  
  # Table data
  writeData(wb, sheet, table,
            startCol = col_start, startRow = row_start + 1)
  
  # Format table data and header
  
  addStyle(wb, sheet, styles$halign_right,
           rows = row_start + 1, 
           cols = (col_start + table_col_center_align):(col_start + ncol(table)),
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$wrap_text,
           rows = row_start + 1, cols = (col_start + 1):(col_start+ncol(table)),
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$border_bottom_thick,
           rows = c(row_start, row_start+1, row_start+1+table_length), cols = (col_start):(col_start+ncol(table) -1),
           gridExpand = TRUE, stack = TRUE)
  
  # Format Table title and  table data header 
  addStyle(wb, sheet, styles$b_font,
           rows = c(row_start, row_start+1), 
           cols = (col_start):(col_start+ncol(table) - 1),
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$valign_mid,
           rows = c(row_start+1, row_start+1+table_length), 
           cols = (col_start):(col_start+ncol(table) - 1),
           gridExpand = TRUE, stack = TRUE)
  
  decimal_format <- paste0("0.", paste(rep("0", decimal_points), collapse = ""))
  num_fmt <- paste0('[>100]#,###;[=-1]"..";', decimal_format)
  
  addStyle(wb, sheet, createStyle(numFmt = num_fmt),
             rows = (row_start + 2):(row_start+1+table_length), cols = (col_start + 1):(col_start+ncol(table)),
             gridExpand = TRUE, stack = TRUE)
  
  # bold last row is Scotland
  if (hb_table == TRUE){
    addStyle(wb, sheet, styles$b_font,
             rows = row_start+1+table_length, 
             cols = (col_start):(col_start+ncol(table) - 1),
             gridExpand = TRUE, stack = TRUE)
    addStyle(wb, sheet, styles$halign_left,
             rows = (row_start + 1):(row_start+1+table_length), 
             cols = (col_start),
             gridExpand = TRUE, stack = TRUE)
    addStyle(wb, sheet, styles$halign_right,
             rows = (row_start + 1):(row_start+1+table_length), 
             cols = (col_start + 1):(col_start+ncol(table) - 1),
             gridExpand = TRUE, stack = TRUE)
  } else {
    addStyle(wb, sheet, styles$halign_centre,
             rows = (row_start + 1):(row_start+1+table_length), 
             cols = (col_start):(col_start+ncol(table) - 1),
             gridExpand = TRUE, stack = TRUE)
  }
  
  if (tail(names(table),1) == "Scotland"){
    addStyle(wb, sheet, styles$b_font,
             rows = (row_start + 1):(row_start+1+table_length), 
             cols = (col_start+ncol(table) - 1),
             gridExpand = TRUE, stack = TRUE)
  }
  
  footnote_counter <- 0
  
  if (!is.null(footnote)) {
    writeData(wb, sheet, footnote,
              startCol = col_start, 
              startRow = row_start+2+table_length)
    addStyle(wb, sheet, styles$font_sz8,
             rows = row_start+2+table_length, 
             cols = col_start)
    footnote_counter <- footnote_counter + 1
  }
  
  # Write explanation for .. (if any)
  if (table %>%
      select(where(is.double)) %>%
      purrr::map_lgl(~ any(. <0)) %>%
      any == TRUE) {
    
    writeData(wb, sheet, "..  shown where no data is available",
              startRow = row_start+table_length+2 + footnote_counter, 
              startCol = 1)
    footnote_counter <- footnote_counter + 1
  }
  
}

wrt_kpi_figure <- function(wb, sheet,
                           figure_header, figure,
                           col_start, row_start,
                           fig_height, fig_width,
                           styles, 
                           header_width = 8,
                           ...){
  # Graph
  print(figure)
  insertPlot(wb, sheet, 
             width = fig_width, height = fig_height, 
             startRow = row_start + 1, startCol = col_start)
  # Graph Title
  writeData(wb, sheet, figure_header,
            startRow = row_start, 
            startCol = col_start)
  
  # Format Graph Title
  setRowHeights(wb, sheet, rows = row_start, heights = 24)
  mergeCells(wb, sheet, rows = row_start, cols = col_start:(col_start + header_width))
  addStyle(wb, sheet, styles$b_font,
           rows = row_start,
           cols = col_start,
           gridExpand = TRUE, stack = TRUE)
  
  addStyle(wb, sheet, styles$wrap_text,
           rows = row_start, 
           cols = col_start,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet, styles$halign_left,
           rows = row_start, 
           cols = col_start,
           gridExpand = TRUE, stack = TRUE)
  
}

#### Function to write to the KPI data sheet for excel reports - One table ####
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
                             footnote = NULL,...) {
  
  row_counter <- 1
  # Titles
  writeData(wb, sheet, kpi_subt,
            colNames = FALSE, startCol = 1, startRow = row_counter)
  writeData(wb, sheet, scr_period,
            colNames = FALSE, startCol = 1, startRow = row_counter + 1)
  if (grepl("1.1", sheet)) {
    writeData(wb, sheet, "See Notes tab for other look-back periods",
              colNames = FALSE, startCol = 1, startRow = row_counter + 2)
  }
  
  # Format titles
  addStyle(wb, sheet, styles$kpi_title,
           rows = row_counter, cols = 1,
           gridExpand = TRUE, stack = TRUE)
  
  row_counter <- row_counter + 5
  kpi_data <- kpi_data %>% 
    no_data_reporting_ref()
  
  wrt_kpi_table(wb, sheet,
                1, row_counter,
                kpi_table_header, kpi_data,
                styles,
                footnote=footnote,
                ...)
  
  row_counter <- row_counter + nrow(kpi_data)+4
  
  wrt_kpi_figure(wb, sheet,
                 fig_legend, kpi_fig,
                             1, row_counter,
                             4, 10,
                             styles,
                 ...)
  
  
  # Format worksheet
  setColWidths(wb, sheet, cols = 1, widths = 28)
  setColWidths(wb, sheet, cols = 2:ncol(kpi_data), widths = 12)
  
}




#### Function to write to the KPI data sheets for excel reports - 2 tables ####
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
                              footnote_a = FALSE, footnote_b = FALSE,
                              col_width = 15,...) {
  
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
  
  
 
  # Convert any NAs to -1 to assign .. for no data
  kpi_data1 <- kpi_data1 %>% 
    no_data_reporting_ref()
  
  row_counter <- 5
  
  wrt_kpi_table(wb, sheet,
                1, row_counter,
                kpi_table_1yr_header, kpi_data1,
                styles,
                ...)
  
  row_counter <- row_counter + nrow(kpi_data1) + 4
  
  # Convert any NAs to -1 to assign .. for no data
  kpi_data2 <- kpi_data2 %>% 
    no_data_reporting_ref()
  
  wrt_kpi_table(wb, sheet,
                1, row_counter,
                kpi_table_5yr_header, kpi_data2,
                styles,
                ...)
  
  
  row_counter <- row_counter + nrow(kpi_data2) + 4
  
  # Graph
  wrt_kpi_figure(wb, sheet,
                 fig_legend, kpi_fig,
                 1, row_counter,
                 4, 10,
                 styles,
                 header_width = 9,
                 ...)
  
  # Format work sheet
  setColWidths(wb, sheet, cols = 1, widths = 28)
  setColWidths(wb, sheet, cols = 2:ncol(kpi_data1), widths = col_width)

  
}


#### Function to write to the KPI 1 Deprivation analysis sheets for excel reports ####
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
                              kpi_data, kpi_fig1, kpi_fig2, styles,
                              plot_1_width = 7.5, plot_2_width = 7.5, ...) {
  
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

  

  # Convert any NAs to -1 to assign .. for no data
  kpi_data <- kpi_data %>% 
    no_data_reporting_ref()
  
  wrt_kpi_table(wb, sheet,
                1, 5,
                kpi_table_header, kpi_data,
                styles,
                ...)
  
  
  wrt_kpi_figure(wb, sheet,
                 fig_legend1, kpi_fig1,
                 1, 6+nrow(kpi_data)+4,
                 4, plot_1_width,
                 styles, 
                 header_width = 5,
                 ...)
  
  wrt_kpi_figure(wb, sheet,
                 fig_legend2, kpi_fig2,
                 10, 6+nrow(kpi_data)+4,
                 4, plot_1_width,
                 styles, 
                 header_width = 10,
                 ...)
  
  # Format worksheet
  setColWidths(wb, sheet, cols = 1, widths = 19)
  setColWidths(wb, sheet, cols = 2:ncol(kpi_data), widths = 15)

  
}



#### Function to write to the KPI data sheets for excel reports - 3 tables ####
## wb: workbook with a sheet named Notes
## sheet: string with the name of the Sheet to write the data to
## scr_period: String with the Screening period
## kpi_title: String with the KPI title
## kpi_subt: String with the KPI subtitle
## fig_legend: String with the legend for the figure
## kpi_data1: tibble with KPI data, Table 1 
## kpi_data2: tibble with KPI data, Tabel 2 
## kpi_fig: Figure presenting KPI data
## styles: Vector with Excel formatting styles

# wrt_kpi_3x_tables <- function(wb, sheet, scr_period, kpi_subt, fig_legend,
#                               kpi_table_1_header, kpi_table_2_header, 
#                               kpi_table_3_header,  
#                               kpi_table_1, kpi_table_2, kpi_table_3, 
#                               kpi_fig, 
#                               styles, 
#                               footnote_a = FALSE, footnote_b = FALSE) {
#   
#   # Format work sheet
#   setColWidths(wb, sheet, cols = 1, widths = 28)
#   setColWidths(wb, sheet, cols = 2:ncol(kpi_table_1), widths = 7)
#   
#   ## Titles
#   writeData(wb, sheet, kpi_subt,
#             colNames = FALSE, startCol = 1, startRow = 1)
#   writeData(wb, sheet, scr_period,
#             colNames = FALSE, startCol = 1, startRow = 2)
#   if (grepl("1.1", sheet)) {
#     writeData(wb, sheet, "See Notes tab for other look-back periods",
#               colNames = FALSE, startCol = 1, startRow = 3)
#   }
#   
#   #  Title format
#   addStyle(wb, sheet, styles$kpi_title,
#            rows = 1, cols = 1,
#            gridExpand = TRUE, stack = TRUE)
#   
#   row_counter <- 5
#   
#   wrt_kpi_table(wb, sheet,
#                             1, row_counter,
#                 kpi_table_1_header, kpi_table_1,
#                             styles,
#                             table_col_center_align = 1)
#   
#   row_counter <- row_counter + 1 + nrow(kpi_table_1) + 3
#   
#   wrt_kpi_table(wb, sheet,
#                 1, row_counter,
#                 kpi_table_2_header, kpi_table_2,
#                 styles,
#                 table_col_center_align = 1)
#   
#   row_counter <- row_counter + 1 + nrow(kpi_table_2) + 3
#   
#   wrt_kpi_table(wb, sheet,
#                 1, row_counter,
#                 kpi_table_3_header, kpi_table_3,
#                 styles,
#                 table_col_center_align = 1)
#   
#   row_counter <- row_counter + 1 + nrow(kpi_table_3) + 2
#   
#   
#   writeData(wb, sheet, "..  shown where no data is available",
#               startRow = row_counter, 
#               startCol = 1)
#   wrt_kpi_figure(wb, sheet,
#                  fig_legend, kpi_fig,
#                  1, row_counter,
#                  4, 10,
#                  styles)
#   
#   
# 
# 
#   
# }












