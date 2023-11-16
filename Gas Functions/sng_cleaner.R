sng_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "FuelUsedInSNG")
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
  
  fuel_type <- raw_utility |> 
    filter(sheet == "FuelUsedInSNG", 
           address == "C6") |> 
    pull(character) |> 
    str_trim(side = "both")
  
  unit <- raw_utility |> 
    filter(sheet == "FuelUsedInSNG", 
           address == "C7") |> 
    pull(character) |> 
    str_trim(side = "both")
  
  sng <- raw_utility |> 
    filter(sheet == "FuelUsedInSNG", 
           row %in% c(8:14), 
           col %in% c(1:3)) |> 
    mutate(value = case_when(
      data_type == "character" ~ character,
      TRUE ~ content
    )) |> 
    select(col, row, value) |> 
    pivot_wider(values_from = value, names_from = col) |> 
    select(-c(row)) |> 
    rename(`Year Label` = 1, 
           "Year" = 2, 
           "Quantity" = 3) |> 
    mutate(Utility = utility_name,
           utility_id = utility_id,
           `Fuel Type` = na_if(fuel_type, "N/A"), #replacing text "N/A" with true null values
           `Unit` = na_if(unit, "N/A"), #replacing text "N/A" with true null values
           Quantity = round(as.numeric(Quantity)), 
           `Year Label` = str_trim(`Year Label`)) |> 
    relocate(Utility, .before = `Year Label`) |> 
    relocate(utility_id, .before = `Year Label`)
  
  return(sng)
  
}