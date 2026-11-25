################################################################################
## R Cervical Screening Publication - KPI 1-3
## func_cervical_kpi_calc_graph.R
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

#### Function to calculate percentages
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
  
  if (calc_coverage == TRUE) { 
    
    hb_data_3_5 <- df %>% filter({{ filter_expr }}, age <=49) %>%
      mutate(scr_period = if_else(stringr::str_sub(financial_year,1,4)<all_scr5_5_start,scr3_5, scr5_5)) |> 
      group_by(...) %>%
      summarize(denominator1 = sum({{ denominator }}),
                numerator1 = sum(scr_period),
                percentage1 = (sum(scr_period)/sum({{ denominator }})*100)) %>%
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
    
    if (calc_coverage == TRUE) {
      
      sc_data_3_5 <- df %>% filter({{ filter_expr }}, age <=49) %>%
        mutate(scr_period = if_else(stringr::str_sub(financial_year,1,4)<all_scr5_5_start,scr3_5, scr5_5)) |> 
        group_by(...) %>%
        ungroup({{ col1 }}) %>% 
        summarize(denominator1 = sum({{ denominator }}),
                  numerator1 = sum(scr_period),
                  percentage1 = (sum(scr_period)/sum({{ denominator }})*100)) %>%
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


#### Function to convert no data to '..' for reporting in Excel
## df: dataframe where numbers might be missing due to no count data
no_data_reporting_ref <- function(df) {
  df %>%
    mutate(replace(., is.na(.), -1))
}


#### Function to create graph for uptake across health boards and Scotland total
## df: dataframe with uptake data for each health board and Scotland
## data_col: data column in df which contain the data for the graph, i.e. "Uptake"
## threshold: uptake threshold, i.e. 80 for 80% threshold line across graph

# hb_graph <- function(df, data_col, threshold) {
#   
#   scot_sep_colours <- c("#3F3685", "#3F3685", "#3F3685", # colour of bars 1-3
#                         "#3F3685", "#3F3685", "#3F3685", # colour of bars 4-6
#                         "#3F3685", "#3F3685", "#3F3685", # colour of bars 7-9
#                         "#3F3685", "#3F3685", "#3F3685", # colour of bars 10-12
#                         "#3F3685", "#3F3685",            # colour of bars 13-14
#                         "#9B4393")                       # colour of bar 15 (Scotland)
#   
#   level_order <- c(df$`NHS Board of Residence`)
#   
#   plot <- df %>% 
#     ggplot(aes(x = factor(`NHS Board of Residence`, level = level_order), 
#                y = !!sym(data_col))) +
#     geom_col(width = 0.75, fill = scot_sep_colours) +
#     geom_hline(yintercept = threshold, linetype = "dashed",
#                linewidth = 0.6, colour = "#C73918") +
#     annotate("text", y = threshold-35, x = 15.8, 
#              label = paste0("Threshold, ", threshold, "%"),
#              color = "#C73918", size = 3, 
#              angle = 90, hjust = 0.1, vjust = 0) +
#     scale_y_continuous(name = data_col,
#                        limits = c(0, 105), n.breaks = 6) +
#     scale_x_discrete(name = "NHS Board of Residence",
#                      labels = function(x) stringr::str_wrap(x, width = 15)) +
#     theme(axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1),
#           panel.background = element_rect("#ECEBF3"))
#   
# }



#### Function to create graph for uptake across health boards and Scotland total
#### where percentages and threshold are small, i.e. on a 100% graph numbers are not easily readable
## df: dataframe with uptake data for each health board and Scotland
## data_col: data column in df which contain the data for the graph, i.e. "Uptake"
## threshold: uptake threshold, i.e. 80 for 80% threshold line across graph

# hb_graph_low_thrs <- function(df, data_col, threshold) {
#   
#   scot_sep_colours <- c("#3F3685", "#3F3685", "#3F3685", # colour of bars 1-3
#                         "#3F3685", "#3F3685", "#3F3685", # colour of bars 4-6
#                         "#3F3685", "#3F3685", "#3F3685", # colour of bars 7-9
#                         "#3F3685", "#3F3685", "#3F3685", # colour of bars 10-12
#                         "#3F3685", "#3F3685",            # colour of bars 13-14
#                         "#9B4393")                       # colour of bar 15 (Scotland)
#   
#   level_order <- c(df$`NHS Board of Residence`)
#   
#   max_value <- round(max(df[[data_col]]), 0) + 1
#   
#   if (max_value < threshold) {
#     max_value <- threshold+2
#   }
#   
#   plot <- df %>% 
#     ggplot(aes(x = factor(`NHS Board of Residence`, level = level_order), 
#                y = !!sym(data_col))) +
#     geom_col(width = 0.75, fill = scot_sep_colours) +
#     geom_hline(yintercept = threshold, linetype = "dashed",
#                linewidth = 0.6, colour = "#C73918") +
#     annotate("text", y = max_value*0.2, x = 15.8, 
#              label = paste0("Threshold, ", threshold, "%"),
#              color = "#C73918", size = 3, 
#              angle = 90, hjust = 0.1, vjust = 0) +
#     scale_y_continuous(name = data_col,
#                        limits = c(0, max_value), n.breaks = 6) +
#     scale_x_discrete(name = "NHS Board of Residence",
#                      labels = function(x) stringr::str_wrap(x, width = 15)) +
#     theme(axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1),
#           panel.background = element_rect("#ECEBF3"))
#   
# }



#### Function to create graph for uptake across health boards and Scotland total,
#### with a data column for each secondary category, i.e. SIMD
## df: dataframe with uptake data for each health board and Scotland
## data_col: data column in df which contain the data for the graph, i.e. "Uptake"
## colours: vector of colour hex codes
## threshold: uptake threshold, i.e. 80 for 80% threshold line across graph
## scot_only: set to TRUE if only showing Scotland level data

cervical_column_graph <- function(df, data_col, colours, 
                                  threshold = FALSE, pct = TRUE, scot_only = FALSE) {
  
  cat1 <- colnames(df[1])
  cat2 <- colnames(df[2])

  if (grepl("SIMD", cat1) | grepl("Age", cat1)) {
    xaxis_angle <- 0
    xvjust <- 1
    xhjust <- 0.5
  } else {
    xaxis_angle <- 45
    xvjust <- 1
    xhjust <- 1
    cat1_levels <- unique(stringr::str_replace(levels(df[[1]]), " and ", " & "))
    df <- df %>% 
      mutate(!!sym(cat1) := factor(stringr::str_replace(!!sym(cat1), " and ", " & "),levels=cat1_levels))
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
               position = position_dodge()#,
               #width = 0.75
               ) +
      scale_x_discrete(name = {{ cat1 }},
                       labels = function(x) stringr::str_wrap(x, width = 15)
                       )  +
      scale_fill_manual(name = {{cat2}},
                        values = colours,
                        guide = guide_legend(order = 2)
      )
  }
  
  if (threshold != FALSE) {
        plot <- add_threshold_line(plot, threshold)
    }
  
  if (pct == TRUE) {
    plot <- plot +
      scale_y_continuous(name = stringr::str_wrap({{ data_col }}, width = 15),
                         limits = c(0, 105), n.breaks = 6,
                         expand = c(0, 0))

  }
  
  if (pct == FALSE) {
    
    max_value <- ceiling(max(df[[data_col]]))
    
    plot <- plot +
      scale_y_continuous(name = stringr::str_wrap({{ data_col }}, width = 15),
                         limits = c(0, max_value+1), n.breaks = 6,
                         expand = c(0, 0)
                         ) 
  }
  
  plot <- add_phs_plot_theme(plot, xaxis_angle = xaxis_angle, xvjust = xvjust, xhjust = xhjust)
  
  plot
}


add_phs_plot_theme <- function(plot, xaxis_angle = 0, xvjust = 1, xhjust = 0.5){
  plot <- plot +
    theme_bw() +
    theme(axis.text.x = element_text(angle = xaxis_angle, hjust = xhjust, vjust = xvjust),
          panel.border = element_rect(colour = "black", fill=NA, linewidth=1),
          panel.grid.major.x = element_blank(),
          panel.grid.major.y = element_line(colour="#bebebe", linewidth = 0.05),
          panel.grid.minor.y = element_line(colour="#bebebe", linewidth = 0.05),
          axis.title.y = element_text(angle = 0, vjust = 0.5))
  
  plot
}

add_threshold_line <- function(plot, threshold){
  threshold_line <- data.frame(yintercept = threshold, label = paste0("Threshold ", threshold, "%"))
  plot <- plot +
    geom_hline(data = threshold_line, aes(yintercept = yintercept, colour = label), 
               linewidth = 0.75, linetype = "dashed") +
    scale_colour_manual(name = "", values = "#F46A25",
                        guide = guide_legend(order = 1))
  
  plot
}
## Create palette ----

# When used in a plotting function it helps assign which colours will be visible within the plot
# Uses the number of unique values in the colour category to pick groups of colours to assign
# so it is forced for small numbers

create_palette <- function(colour) {
  
  num_required <- length(unique(colour))
  
  if (num_required<=8) {
    options <- c("#12436D", "#28A197", "#801650", #"#F46A25", 
                 "#3F085C", "#3E8ECC", "#3D3D3D", "#A285D1")
    
  } else {
    
    options <- c("#12436D", "#94AABD", "#28A197", "#B4DEDB", "#801650", "#CCA2B9",
                 #"#F46A25", 
                 "#FBC3A8", "#3D3D3D", "#A8A8A8", "#3E8ECC", "#A8CCE8",
                 "#3F085C", "#A285D1")
  }
  
  palette <- options[1:num_required]
  return(palette)
  
}

#### Function to create line graph with optional threshold line, annotation and
#### points
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
    scale_y_continuous(name = stringr::str_wrap({{ data_col }}, width = 15),
                       limits=c(0,100), expand=c(0,0)) +
    scale_x_discrete(expand=c(0.05,0.05))   
  
  # Optionally add annotation for 2020/21
  if (annotate==TRUE){
    if ("2020/21" %in% pull(df, {{ data_col }})){
    plot <- plot +
      geom_vline(xintercept = "2020/21", linewidth = 0.3, linetype="dashed") +
      annotate("label", label="A", x="2020/21", y=95, size=3)
    }
  }
  
  # Optionally add threshold line
  if (threshold != FALSE){
    plot <- add_threshold_line(plot, threshold)
  }
  
  # Optionally add points to line
  if (points==TRUE){
    plot <- plot +
      geom_point(aes(x=!!sym(names(df[2])), y=!!sym(data_col)), colour=colours)
  }
  
  # Add main line to plot (last to appear over other lines)
  plot <- plot +
    geom_line(colour=colours)
  
  # Add phs theme
    plot <- add_phs_plot_theme(plot)
  
  plot
  
  return(plot)
}

create_fin_year <- function(x){
    y <- lubridate::year(x %m-% months(3))
    
    ifelse(
      is.na(x),
      NA_character_,
      paste0(y, "/", substr(y + 1, 3, 4))
    )
}

add_age <- function(df){
  df |> 
    mutate(age = age_calculate(dob, year_end + 1)) %>%
    mutate(
      age_group = case_when(
        age <= 24 ~ "Under 25",
        age >= 25 & age <= 29 ~ "25-29",
        age >= 30 & age <= 34 ~ "30-34",
        age >= 35 & age <= 39 ~ "35-39",
        age >= 40 & age <= 44 ~ "40-44",
        age >= 45 & age <= 49 ~ "45-49",
        age >= 50 & age <= 54 ~ "50-54",
        age >= 55 & age <= 59 ~ "55-59",
        age >= 60 & age <= 64 ~ "60-64",
        age >= 65 ~ "65+"
      ),
      age_group_25_64 = case_when(
        age >= 25 & age <= 64 ~ "25-64",
        age <= 24 ~ "Under 25",
        age >= 65 ~ "65+"
      ),
      age_group2549_5064 = case_when(
        age <= 24 ~ "Under 25",
        age >= 25 & age <= 49 ~ "25-49",
        age >= 50 & age <= 64 ~ "50-64",
        age >= 65 ~ "65+"
      )
    )
}

add_flags <- function(df, invite = TRUE, exclusion = TRUE, uptake = TRUE){
  
  if (invite){
    df <- df |> 
      mutate(invite_fy_flag = if_else(age >= 25 & # Remove any <25 from invite_fy_flag as there are very few
                                        !is.na(dateof_prompt), 1, 0))
  }
  
  if (exclusion){
    df <-  df |>      
      mutate(exclusion_flag = if_else(type_of_exclusion %in% flagged_exclusions, 1, 0))
  }
  
  if (uptake){
    df <- df |> 
      # Screened within 6 month of a prompt (Invite or Reminder)
      mutate(uptake_flag = if_else(!is.na(dateof_examination), 1, 0))
  }
  
  df
}

add_postcode <- function(df){
  df |>
    mutate(pc_invite = if_else(invite_fy_flag == 1,
                                 coalesce(pc8_invite, pc8_demo),
                                 pc8_demo)) %>%
      mutate(postcode = if_else(uptake_flag == 1,
                                coalesce(pc8_screen, pc_invite),
                                pc_invite))  |>
      left_join(simd, by = join_by(postcode == pc8)) %>%
      mutate(
        hb2019name = stringr::str_replace(hb2019name, "NHS ", ""),
        simd2020v2_sc_quintile = factor(simd2020v2_sc_quintile,
                                        levels = c("1", "2", "3", "4", "5")
        )
      ) %>%
      mutate(gp_invite = if_else(invite_fy_flag == 1,
                                 coalesce(practice_invite, current_practice),
                                 current_practice)) %>%
      mutate(gp_practice_code = if_else(uptake_flag == 1,
                               coalesce(practice_res, gp_invite),
                               gp_invite)) %>%
      mutate(
        gp_practice_code = if_else(is.na(gp_practice_code), "NOREG", gp_practice_code),
        PracticeCode = substr(gp_practice_code, 2, 6),
        PracticeCode = if_else(PracticeCode == "OREG" | is.na(PracticeCode), "NOREG", PracticeCode)
      ) |> 
    select(!c(gp_invite, pc_invite))
}

calculate_percent <- function(df, num, den, ...){
  df |> 
    group_by(...) |> 
    summarise(
      numerator = sum({{ num }}, na.rm = TRUE),
      denominator = sum({{ den }}, na.rm = TRUE),
      percentage = numerator/denominator * 100
      ) |>
    ungroup() |> 
    complete(...)
}

calculate_kpi <- function(df, num, den, ..., 
                          hb = hb2019, financial_year = financial_year,
                          scot = TRUE){
  df <- df |> 
    filter({{ den }} == 1)
  
  output <- df |> 
    calculate_percent(num = {{ num}}, den = {{ den }}, 
                      {{ hb }}, {{ financial_year }}, ...)
  
  if (scot){
    scot_level <- df |> 
      calculate_percent(num = {{ num}}, den = {{ den }}, 
                        {{ financial_year }}, ...) |> 
      mutate({{ hb }} := "S92000003")
    
    output <- output |> 
      bind_rows(scot_level)
  }
  
  
}

#### End of Script ----
