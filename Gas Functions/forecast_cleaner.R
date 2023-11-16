forecast_cleaner <- function(directory, workbook){
  
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "SalesByCategory_Large")
  )
  
  sales_check <- raw_utility |> 
    filter(sheet == "SalesByCategory_Large", 
           address == "J11") |> 
    pull(content)
  
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
  
  if(sales_check == "0"){
    
    return(NULL)
    
  }
  
  else{
    
    raw_sales <- filter(raw_utility, sheet %in% c("SalesByCategory_Large"))
    
    sales_categories <- raw_sales |> 
      filter(row == 10 & col %in% c(1:9)) |> 
      pull(character) |> 
      str_replace_all("\\\r\\n", " ") |> 
      str_replace("Total Annual Gas Sales to Ultimate Consumers in Minnesota", "Total")
    
    sales_categories[1] <- "forecast_year"
    sales_categories[2] <- "category"
    
    sales_data <- raw_sales |> 
      filter(row %in% c(11:37) & col %in% c(1:9)) |> 
      mutate(content = case_when(
        col == 2 ~ character, 
        TRUE ~ content
      )) |> 
      select(row, col, content) |> 
      pivot_wider(names_from = col, values_from = content) |> 
      fill(`1`) |> 
      select(-c(row)) |> 
      mutate(`2` = str_replace(`2`, "\\*", ""), 
             `2` = str_replace(`2`, "MCF Transp.", "Mcf Transportation"), 
             `2` = str_replace(`2`, "MCF Sales", "Mcf Sales")) 
    
    colnames(sales_data) <- sales_categories
    
    sales_data <- sales_data |> 
      pivot_longer(!c(forecast_year, category), names_to = "customer_category", values_to = "value") |> 
      pivot_wider(names_from = "category", values_from = "value") |> 
      mutate(Utility = utility_name, 
             utility_id = utility_id,
             `Report Year` = report_year) |> 
      relocate(Utility, .before = forecast_year) |>
      relocate(utility_id, .before = forecast_year) |> 
      relocate(`Report Year`, .before = customer_category) |> 
      rename(`Customer Category` = customer_category, 
             `Forecast Year` = forecast_year) |> 
      mutate(Customers = round(as.numeric(Customers)), 
             `Mcf Sales` = round(as.numeric(`Mcf Sales`)), 
             `Mcf Transportation` = round(as.numeric(`Mcf Transportation`)))
    
    return(sales_data)
    
  }
  
  
}