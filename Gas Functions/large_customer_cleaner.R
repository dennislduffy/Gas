
large_customer_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"), 
    sheets = c("Registration", "CommInd200MCF")
  )
  
  utility_id <- raw_utility |> 
    filter(sheet == "Registration", 
           address == "C5") |> 
    pull(content)
  
  utility_name <- raw_utility |> 
    filter(sheet == "Registration", 
           row == 9 & col == 3) |> 
    pull(character)
  
  report_year = raw_utility |> 
    filter(sheet == "Registration", 
           row == 6 & col == 3) |> 
    pull(content)
  
  large_cust_table <- raw_utility |> 
    filter(sheet == "CommInd200MCF", 
           col %in% c(4:5) & row %in% c(8:10)) |>
    mutate(value = case_when(
      data_type == "character" ~ character, 
      TRUE ~ content
    )) |> 
    select(row, col, value) |> 
    pivot_wider(names_from = "col", values_from = "value") |> 
    select(-c(row)) |> 
    row_to_names(row_number = 1) |> 
    rename(`Commercial / Industrial > 200 Mcf on Peak Day` = 1, 
           `Commercial / Industrial Interruptible > 200 Mcf on Peak Day` = 2)
  
  large_cust_table <- raw_utility |> 
    filter(sheet == "CommInd200MCF", 
           col == 3 & row %in% c(8:10)) |> 
    select(row, col, character) |> 
    mutate(character = case_when(
      row == 8 ~ "units", 
      row == 9 ~ "Number of Customers", 
      TRUE ~ character
    )) |> 
    select(character) |> 
    row_to_names(row_number = 1) |> 
    bind_cols(large_cust_table) |> 
    pivot_longer(!units, names_to = "customer_type", values_to = "values") |> 
    pivot_wider(names_from = units, values_from = values) |> 
    mutate(`Number of Customers` = as.numeric(`Number of Customers`), 
           `Mcf` = as.numeric(Mcf)) |> 
    mutate(Utility = utility_name, 
           utility_id = utility_id,
           Year = report_year) |> 
    relocate(Utility, .before = customer_type) |> 
    relocate(utility_id, .before = customer_type) |> 
    relocate(Year, .before = customer_type) |> 
    rename(`Customer Type` = customer_type)
  
  return(large_cust_table)
  
}