##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## R Cervical Screening - KPIs for Publication
## func_cervical_kpi_calc_graph.R
## 
## Adult Screening Team
## Created: 25/07/2024
## Last update: 09/07/2025
## 
## Written/run on RStudio Posit Workbench
## Platform: x86_64-pc-linux-gnu
## R version 4.4.2 (Pile of Leaves) - 2024-10-31
##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


# Function to set eligibility flag depending on reporting (financial) year ----
## df: KPI 1.1 data frame for adding eligibility flag to
## fy_yr_start: The year of the start of the reporting (financial) year

cervical_eli_flag <- function(df, 
                              fy_yr_start = yr_start) {
  # Eligible age changed from 20-60 in 2015/16 to 25-64 in 2016/17
  # 2016/17 eligible age: 21-64
  # 2017/18 eligible age: 22-64
  # 2018/19 eligible age: 23-64
  # 2019/20 eligible age: 24-64
  if (fy_yr_start < 2016) {
    df_eli <- df %>%
      mutate(eligible = ifelse(age >=20 & age <=60 & exclusion_flag == 0, 1, 0))
    }
  if (fy_yr_start == 2016) {
    df_eli <- df %>% 
      mutate(eligible = ifelse(age >=21 & age <=64 & exclusion_flag == 0, 1, 0))
  }
  if (fy_yr_start == 2017) {
    df_eli <- df %>% 
      mutate(eligible = ifelse(age >=22 & age <=64 & exclusion_flag == 0, 1, 0))
  }
  if (fy_yr_start == 2018) {
    df_eli <- df %>%
      mutate(eligible = ifelse(age >=23 & age <=64 & exclusion_flag == 0, 1, 0))
  }
  if (fy_yr_start == 2019) {
    df_eli <- df %>% 
      mutate(eligible = ifelse(age >=24 & age <=64 & exclusion_flag == 0, 1, 0))
  }
  if (fy_yr_start >= 2020) {
    df_eli <- df %>% 
      mutate(eligible = ifelse(age >=25 & age <=64 & exclusion_flag == 0, 1, 0))
  }
  
  df_eli
}



# Function to convert no data to '..' for reporting in Excel ----
## df: dataframe where numbers might be missing due to no count data
no_data_reporting_ref <- function(df) {
  df %>% 
    mutate(across(where(is.double), ~ round(., 1))) %>% 
    mutate(replace(., is.na(.), -1))
}



# Function to create column bar graph ----
# for uptake across health boards and Scotland total, with a data column for each secondary category, i.e. SIMD
## df: dataframe with uptake data for each health board and Scotland
## data_col: data column in df which contain the data for the graph, i.e. "Uptake"
## colours: vector of colour hex codes
## threshold: uptake threshold, i.e. 80 for 80% threshold line across graph
## scot_only: set to TRUE if only showing Scotland level data

cervical_column_graph <- function(df, data_col, colours, 
                                  threshold = FALSE, pct = TRUE, scot_only = FALSE) {
  
  cat1 <- colnames(df[1])
  cat2 <- colnames(df[2])

  if (grepl("SIMD", cat1) | grepl("Age", cat1) | scot_only == TRUE) {
    xaxis_angle <- 0
  } else {
    xaxis_angle <- 45
  }
    
  if (scot_only == TRUE) {
    
    plot <- df %>% 
      ggplot(aes(x = !!sym(cat2), 
                 y = !!sym(data_col))) +
      geom_col(width = 0.75, 
               fill = colours) +
      scale_x_discrete(name = {{ cat2 }},
                       labels = function(x) stringr::str_wrap(x, width = 15)) 
    
  } else {
    
    plot <- df %>% 
      ggplot(aes(x = !!sym(cat1), 
                 y = !!sym(data_col))) +
      geom_col(aes(fill = !!sym(cat2)),
               position = position_dodge(),
               width = 0.75) +
      scale_x_discrete(name = {{ cat1 }},
                       labels = function(x) stringr::str_wrap(x, width = 15)) +
      scale_fill_manual(name = {{ cat2 }},
                        values = colours) 
  }
  
    if (threshold != FALSE) {
      plot <- plot +
        geom_hline(yintercept = threshold, linetype = "dashed",
                   linewidth = 0.6, colour = "#A01E25") +
        annotate("text", y = threshold, x = 0.3, 
                 label = paste0("Threshold, ", threshold, "%"),
                 color = "#A01E25", size = 3, 
                 angle = 90, hjust = 1.1, vjust = 1)
    }
  
  if (pct == TRUE) {
    plot <- plot +
      scale_y_continuous(name = data_col,
                         limits = c(0, 105), n.breaks = 6) +
      theme(axis.text.x = element_text(angle = xaxis_angle, hjust = 1, vjust = 1),
            panel.background = element_rect("#ECEBF3"))
  }
  
  if (pct == FALSE) {
    
    max_value <- ceiling(max(df[[data_col]]))
    
    plot <- plot +
      scale_y_continuous(name = data_col,
                         limits = c(0, max_value+1), n.breaks = 6
                         ) +
      theme(axis.text.x = element_text(angle = xaxis_angle, hjust = 1, vjust = 1),
            panel.background = element_rect("#ECEBF3"))
  }
  plot
}

# Function to create line graph ----
# with optional threshold line, annotation and points
## df: dataframe - group data in column 1, x-axis data in column 2 and y-axis
## in column 3
## data_col: string - data column name df for y-axis
## colours: vector - colour hex codes
## threshold: numeric - i.e. 80 for 80% threshold line across graph
## annotate: logical - add annotation line in 2020/21
## points: logical - add points to plot
cervical_line_graph <- function(df, data_col, colours, threshold=FALSE, 
                                annotate=TRUE, points=TRUE){
  
  #Create plot background
  plot <- df %>% 
    ggplot( aes(x=!!sym(names(df[2])), y=!!sym(data_col), group=names(df[1])) ) +
    scale_y_continuous(limits=c(0,100), expand=c(0,0)) +
    scale_x_discrete(expand=c(0.005,0.005)) +
    theme(panel.grid.major.x = element_blank(),
          panel.grid.major.y = element_line(colour = "light grey"),
          panel.grid.minor.y = element_line(colour = "light grey"),
          panel.background = element_rect(fill = "white", colour="light grey" ),
          plot.margin = margin(10, 20, 10, 10))    
  
  # Optionally add annotation for 2020/21
  if (annotate==TRUE){
    plot <- plot +
      geom_vline(xintercept = "2020/21",linewidth = 0.3,linetype="dashed") +
      annotate("label", label="A", x="2020/21", y=95, size=3)
  }
  
  # Optionally add threshold line
  if (threshold != FALSE){
    plot <- plot + 
      geom_hline(yintercept = threshold, linetype = "dashed",linewidth = 0.3, 
                 colour = "#A01E25") +
      annotate("text", y = threshold, x=1, 
               label = paste0("Threshold ",threshold,"%"), 
               color = "#A01E25", size = 3, hjust = 0.05, vjust = -0.2)
  }
  
  # Optionally add points to line
  if (points==TRUE){
    plot <- plot +
      geom_point(aes(x=!!sym(names(df[2])), y=!!sym(data_col)), colour=colours)
  }
  
  # Add main line to plot (last to appear over other lines)
  plot <- plot +
    geom_line(colour=colours)
  
  
  return(plot)
}

