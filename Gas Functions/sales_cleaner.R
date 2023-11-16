sales_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "SalesByCategory_Small", "SalesByCategory_Large")
  )
  
  sales_check <- raw_utility |> 
    filter(sheet == "SalesByCategory_Small", 
           address == "B23") |> 
    pull(is_blank)
  
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
  
  if(sales_check == TRUE){
    
    raw_sales <- filter(raw_utility, sheet %in% c("SalesByCategory_Large"))
    
    sales_categories <- raw_sales |> 
      filter(row == 10 & col %in% c(2:9)) |> 
      pull(character) |> 
      str_replace_all("\\\r\\n", " ") |> 
      str_replace("Total Annual Gas Sales to Ultimate Consumers in Minnesota", "Total")
    
    sales_categories[1] <- "category"
    
    sales_data <- raw_sales |> 
      filter(row %in% c(11:13) & col %in% c(2:9)) |> 
      mutate(content = case_when(
        col == 2 ~ character, 
        col != 2 & data_type == "numeric" ~ content, 
        TRUE ~ NA
      )) |> 
      select(row, col, content) |> 
      pivot_wider(names_from = col, values_from = content) |> 
      select(-c(row)) |> 
      mutate(`2` = str_replace(`2`, "\\*", ""), 
             `2` = str_replace(`2`, "MCF Transp.", "Mcf Transportation"), 
             `2` = str_replace(`2`, "MCF Sales", "Mcf Sales")) 
    
    colnames(sales_data) <- sales_categories
    
    sales_data <- sales_data |> 
      pivot_longer(!category, names_to = "customer_category", values_to = "value") |> 
      pivot_wider(names_from = "category", values_from = "value") |> 
      mutate(Utility = utility_name, 
             utility_id = utility_id,
             `Report Year` = report_year) |> 
      relocate(Utility, .before = customer_category) |> 
      relocate(utility_id, .before = customer_category) |> 
      relocate(`Report Year`, .before = customer_category) |> 
      rename(`Customer Category` = customer_category) |> 
      mutate(Customers = as.numeric(Customers), 
             `Mcf Sales` = round(as.numeric(`Mcf Sales`)), 
             `Mcf Transportation` = round(as.numeric(`Mcf Transportation`)))
    
  }
  
  else{
    
    raw_sales <- filter(raw_utility, sheet == "SalesByCategory_Small")
    
    sales_categories <- raw_sales |>
      filter(row == 22 & col %in% c(1:8)) |> 
      mutate(content = case_when(
        col == 1 ~ content, 
        TRUE ~ character
      )) |> 
      pull(content) |> 
      str_replace_all("\\\r\\\n", " ")
    
    sales_data <- raw_sales |> 
      filter(row %in% c(23:25) & col %in% c(1:8)) |>
      mutate(content = case_when(
        col == 1 ~ character, 
        col != 1 & data_type == "numeric" ~ content, 
        TRUE ~ NA
      )) |> 
      select(row, col, content) |> 
      pivot_wider(names_from = col, values_from = content) |> 
      select(-c(row)) |> 
      mutate(`1` = str_trim(str_replace_all(`1`, "\\*", ""), side = "both")) |> 
      mutate(`1` = str_replace(`1`, " \\(at year's end\\)", ""))
    
    colnames(sales_data) <- sales_categories
    
    sales_data <- sales_data |> 
      pivot_longer(!1, names_to = "customer_category", values_to = "value") |> 
      pivot_wider(names_from = 1, values_from = value) |> 
      mutate(Utility = utility_name, 
             utility_id = utility_id,
             `Report Year` = report_year) |> 
      relocate(Utility, .before = customer_category) |> 
      relocate(utility_id, .before = customer_category) |> 
      relocate(`Report Year`, .before = customer_category) |> 
      rename(`Customer Category` = 4, 
             Customers = 5) |> 
      mutate(Customers = as.numeric(Customers), 
             `Mcf Sales` = round(as.numeric(`Mcf Sales`)), 
             `Mcf Transportation` = round(as.numeric(`Mcf Transportation`)))
    
    
  }
  
  return(sales_data)
  
}