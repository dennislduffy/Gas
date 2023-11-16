sales_revenue_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "SalesRevenue")
  )
  
  utility_name <- raw_utility |> 
    filter(sheet == "Registration", 
           row == 9 & col == 3) |> 
    pull(character)
  
  report_year = raw_utility |> 
    filter(sheet == "Registration", 
           row == 6 & col == 3) |> 
    pull(content)
  
  sales_revenue <- raw_utility |> 
    filter(sheet == "SalesRevenue", 
           col %in% c(1:7) & row %in% c(7:9)) |> 
    mutate(value = case_when(
      data_type == "character" ~ character, 
      TRUE ~ content
    )) |> 
    select(row, col, value) |> 
    mutate(value = str_replace_all(value, "\\r\\n", " "), 
           value = str_replace_all(value, "\\*", ""), 
           value = str_trim(value, side = "both")) |> 
    pivot_wider(names_from = "col", values_from = "value") |> 
    select(-c(row)) |> 
    row_to_names(row_number = 1) |> 
    pivot_longer(!1, names_to = "customer_type", values_to = "value") |> 
    pivot_wider(names_from = 1, values_from = "value") |> 
    mutate(across(.cols = c(SALES, TRANSPORTATION), .fns = as.numeric), 
           Utility = utility_name, 
           Year = report_year) |> 
    relocate(Utility, .before = customer_type) |> 
    relocate(Year, .before = customer_type) |> 
    rename(`Customer Type` = customer_type, 
           Sales = SALES, 
           Transportation = TRANSPORTATION)
  
  return(sales_revenue)
  
}
