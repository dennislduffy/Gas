peak_day_forecast_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "PeakDayForecast")
  )
  
  utility_name <- raw_utility |> 
    filter(sheet == "Registration", 
           row == 9 & col == 3) |> 
    pull(character)
  
  utility_id <- raw_utility |> 
    filter(sheet == "Registration", 
           address == "C5") |> 
    pull(content)
  
  report_year = raw_utility |> 
    filter(sheet == "Registration", 
           row == 6 & col == 3) |> 
    pull(content)
  
  column_names <- raw_utility |>
    filter(sheet == "PeakDayForecast",
           row == 17 & col %in% c(1:10)) |>
    pull(character) |>
    replace_na("year_label") |> 
    str_replace_all("\\r\\n", " ") |> 
    str_replace("CONSUMPTION BY CATEGORY \\-PEAK\\-DAY\\- ", "row") 
  
  customers <- raw_utility |>
    filter(sheet == "PeakDayForecast",
           col %in% c(1, 3:9, 11) & row %in% c(18:35)) |>
    mutate(content = case_when(
      data_type == "character" ~ character,
      TRUE ~ content
    )) |>
    select(row, col, content) |>
    pivot_wider(names_from = "col", values_from = "content") |>
    filter(row %% 2 == 0) 
  
  colnames(customers) <- column_names
  
  customers <- customers|> 
    select(-c(row)) |> 
    pivot_longer(!year_label, names_to = "customer_type", values_to = "customer_count") |> 
    mutate(customer_type = str_replace(customer_type, "\\(should equal Column 8, on Basic Forecast worksheet tab\\)", ""), 
           customer_type = str_replace(customer_type, "\\*", "")) |> 
    separate(customer_type, sep = " McF ", c("customer_type", "rule_section")) |> 
    mutate(customer_count = as.numeric(customer_count))
  
  gas_sales <- raw_utility |> 
    filter(sheet == "PeakDayForecast", 
           col %in% c(1, 3:9, 11) & row %in% c(18:35)) |> 
    mutate(content = case_when(
      data_type == "character" ~ character,
      TRUE ~ content
    )) |>
    select(row, col, content) |>
    pivot_wider(names_from = "col", values_from = "content") |>
    filter(row %% 2 != 0)
  
  column_names <- raw_utility |>
    filter(sheet == "PeakDayForecast",
           row == 17 & col %in% c(1:10)) |>
    pull(character) |>
    replace_na("forecast_year") |> 
    str_replace_all("\\r\\n", " ") |> 
    str_replace("CONSUMPTION BY CATEGORY \\-PEAK\\-DAY\\- ", "row") 
  
  colnames(gas_sales) <- column_names
  
  peak_day <- gas_sales |> 
    select(-c(row)) |> 
    pivot_longer(!forecast_year, names_to = "customer_type", values_to = "Mcf") |> 
    select(-c(customer_type)) |> 
    bind_cols(customers) |> 
    select(year_label, forecast_year, customer_type, customer_count, Mcf, rule_section) |> 
    mutate(Utility = utility_name, 
           utility_id = utility_id,
           `Report Year` = report_year, 
           Mcf = as.numeric(Mcf)) |> 
    relocate(Utility, .before = year_label) |> 
    relocate(utility_id, .before = year_label) |> 
    relocate(`Report Year`, .before = year_label) |> 
    rename(`Year Label` = year_label, 
           `Forecast Year` = forecast_year, 
           `Customer Type` = customer_type, 
           `Customer Count` = customer_count, 
           `Rule Section` = rule_section)
  
  return(peak_day)
  
}
