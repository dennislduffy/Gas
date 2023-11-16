basic_forecast_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "BasicForecast")
  )
  
  report_year <- raw_utility |> 
    filter(sheet == "Registration", 
           row == 6, col == 3) |> 
    pull(numeric)
  
  utility_name <- raw_utility |> 
    filter(sheet == "Registration", 
           row == 9 & col == 3) |> 
    pull(character)
  
  utility_id <- raw_utility |> 
    filter(sheet == "Registration", 
           address == "C5") |> 
    pull(content)
  
  first_table_columns <- raw_utility |> 
    filter(sheet == "BasicForecast", 
           row == 8 & col %in% c(1:8, 10:11)) |> 
    pull(character) |> 
    replace(1, "Year Label") |> 
    replace(2, "Year") |> 
    str_replace_all("\\r\\n", " ") |> 
    str_replace("Columns 1-5", "Columns 3-7") |> 
    str_replace("Columns 1-6", "Columns 3-8") |> 
    str_replace("\\*", "")
  
  first_table_values <- raw_utility |> 
    filter(sheet == "BasicForecast", 
           row %in% c(9:15) & col %in% c(1:8, 10:11)) |>
    mutate(value = case_when(
      col == 1 ~ character, 
      TRUE ~ content
    )) |> 
    select(row, col, value) |> 
    pivot_wider(names_from = col, values_from = value) |> 
    select(-c(row)) |> 
    mutate(across(-c(`1`), as.numeric)) |> 
    mutate(across(-c(`1`), round))
  
  colnames(first_table_values) <- first_table_columns
  
  first_table_values <- first_table_values |> 
    mutate(Utility = utility_name, 
           utility_id = utility_id,
           `Report Year` = report_year) |> 
    relocate(Utility, .before = `Year Label`) |> 
    relocate(utility_id, .before = `Year Label`) |> 
    relocate(`Report Year`, .before = `Year Label`)
  
  second_table_columns <- raw_utility |> 
    filter(sheet == "BasicForecast", 
           row == 22 & col %in% c(3:8, 10:13)) |> 
    pull(character) |> 
    str_replace_all("\\r\\n", " ") |> 
    str_replace("Sum Columns 11-14", "Sum Columns 15-18") |> 
    str_replace("\\*", "") |> 
    str_replace("; Should Equal Column 7 ", "")
  
  second_table_values <- raw_utility |> 
    filter(sheet == "BasicForecast", 
           row %in% c(23:29) & col %in% c(3:8, 10:13)) |> 
    mutate(value = case_when(
      col == 1 ~ character, 
      TRUE ~ content
    )) |> 
    select(row, col, value) |> 
    pivot_wider(names_from = col, values_from = value) |> 
    select(-c(row)) |> 
    mutate(across(.cols = everything(), as.numeric)) |> 
    mutate(across(.cols = everything(), round))
  
  colnames(second_table_values) <- second_table_columns
  
  basic_forecast <- first_table_values |> 
    bind_cols(second_table_values)
  
  return(basic_forecast)
}