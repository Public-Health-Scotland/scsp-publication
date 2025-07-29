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


# Function to calculate percentages ----
## df: dataframe to analyse uptake for
## filter_expr: add expression for filtering data, use brackets around the expression
## numerator: add the name of the column to be the numerator
## denominator: add the name of the column to be the denominator
## ...: Add health board variable, and any number of additional group_by groups to group the data
## scot: default TRUE, change to FALSE for no Scotland total
## calc_coverage: default FALSE, if KPI 1.1 Coverage is calculated set to TRUE
## fy_yr_start: The year of the start of the reporting (financial) year
## all_scr5_5_start_yr: The year when the 5.5 year look back period should be implemented to the KPI analysis

cervical_pct <- function(df, filter_expr, numerator, denominator, ..., 
                         scot = TRUE,
                         calc_coverage = FALSE,
                         fy_yr_start = yr_start, 
                         all_scr5_5_start_yr = all_scr5_5_start) {
  
  if (fy_yr_start < all_scr5_5_start_yr & calc_coverage == TRUE) { 
    
    hb_data_3_5 <- df %>% filter({{ filter_expr }}, age <=49) %>%
      group_by(...) %>%
      summarize(denominator1 = sum({{ denominator }}),
                numerator1 = sum(scr3_5),
                percentage1 = (sum(scr3_5)/sum({{ denominator }})*100)) %>%
      ungroup()
    
    hb_data_5_5 <- df %>% filter({{ filter_expr }}, age >=50, age <=64) %>%
      group_by(...) %>%
      summarize(denominator1 = sum({{ denominator }}),
                numerator1 = sum({{ numerator }}),
                percentage1 = (sum({{ numerator }})/sum({{ denominator }})*100)) %>%
      ungroup()
    
    hb_data <- rbind(hb_data_3_5, hb_data_5_5) %>% 
      group_by(...) %>%
      summarize(Denominator = sum(denominator1),
                Numerator = sum(numerator1),
                Percentage = (sum(numerator1)/sum(denominator1)*100)) %>%
      ungroup()
    
  } else {
    
    hb_data <- df %>% filter({{ filter_expr }}) %>%
      group_by(...) %>%
      summarize(Denominator = sum({{ denominator }}),
                Numerator = sum({{ numerator }}),
                Percentage = (sum({{ numerator }})/sum({{ denominator }})*100)) %>%
      ungroup()
    
  }
  
  if (scot == TRUE) {
    
    col1 <- colnames(hb_data[1])
    
    if (fy_yr_start < all_scr5_5_start_yr & calc_coverage == TRUE) {
      
      sc_data_3_5 <- df %>% filter({{ filter_expr }}, age <=49) %>%
        group_by(...) %>%
        ungroup({{ col1 }}) %>% 
        summarize(denominator1 = sum({{ denominator }}),
                  numerator1 = sum(scr3_5),
                  percentage1 = (sum(scr3_5)/sum({{ denominator }})*100)) %>%
        ungroup() %>% 
        mutate({{ col1 }} := "S92000003", .before = 1)
      
      sc_data_5_5 <- df %>% filter({{ filter_expr }}, age >=50, age <=64) %>%
        group_by(...) %>%
        ungroup({{ col1 }}) %>% 
        summarize(denominator1 = sum({{ denominator }}),
                  numerator1 = sum({{ numerator }}),
                  percentage1 = (sum({{ numerator }})/sum({{ denominator }})*100)) %>%
        ungroup() %>% 
        mutate({{ col1 }} := "S92000003", .before = 1)
      
      scotland_data <- rbind(sc_data_3_5, sc_data_5_5) %>% 
        group_by(...) %>%
        ungroup({{ col1 }}) %>% 
        summarize(Denominator = sum(denominator1),
                  Numerator = sum(numerator1),
                  Percentage = (sum(numerator1)/sum(denominator1)*100)) %>%
        ungroup() %>% 
        mutate({{ col1 }} := "S92000003", .before = 1)
      
    } else {
      
      scotland_data <-  df %>% filter({{ filter_expr }}) %>% 
        group_by(...) %>%
        ungroup({{ col1 }}) %>% 
        summarize(Denominator = sum({{ denominator }}),
                  Numerator = sum({{ numerator }}),
                  Percentage = (sum({{ numerator }})/sum({{ denominator }})*100)) %>%
        ungroup() %>% 
        mutate({{ col1 }} := "S92000003", .before = 1)
      
    }
    
    total_data <- rbind(hb_data, scotland_data)
    
    col1 <- colnames(total_data[1])
    
    total_data[[{{ col1 }}]] <- factor(total_data[[{{ col1 }}]],
                                       levels = hb_levels)
    
    total_data %>% 
      complete(..., explicit = FALSE)
    
  } else {
    
    hb_data
    
  }
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


# Function to create palette ----
## When used in a plotting function it helps assign which colours will be visible within the plot
## Uses the number of unique values in the colour category to pick groups of colours to assign
## so it is forced for small numbers

create_palette <- function(colour) {
  
  vector <- as.character(colour)
  
  if (length(unique(vector)) ==1 ) {
    palette <- "#12436D"
    
  } else if (length(unique(vector)) ==2) {
    palette <- c("#12436D", "#28A197")
    
  } else if (length(unique(vector)) ==3) {
    palette <- c("#12436D", "#28A197", "#801650")
    
  } else if (length(unique(vector)) ==4) {
    palette <- c("#12436D", "#28A197", "#801650", "#F46A25")
    
  } else if (length(unique(vector)) ==5) {
    palette <- c("#12436D", "#28A197", "#801650", "#F46A25", "#3F085C")
    
    
  } else if (length(unique(vector)) ==6) {
    palette <- c("#12436D", "#28A197", "#801650", "#F46A25", "#3F085C", "#3E8ECC")
    
  } else if (length(unique(vector)) ==7) {
    palette <- c("#12436D", "#28A197", "#801650", "#F46A25", "#3F085C", "#3E8ECC", "#3D3D3D")
    
    
  } else if (length(unique(vector)) ==8) {
    palette <- c("#12436D", "#28A197", "#801650", "#F46A25", "#3F085C", "#3E8ECC", "#3D3D3D", "#A285D1")
    
  } else {
    
    palette <- c("#12436D", "#94AABD", "#28A197", "#B4DEDB", "#801650", "#CCA2B9",
                 "#F46A25", "#FBC3A8", "#3D3D3D", "#A8A8A8", "#3E8ECC", "#A8CCE8",
                 "#3F085C", "#A285D1")
  }
  
  return(palette)
  
}


