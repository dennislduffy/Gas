design_day_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "DesignDay")
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
  
  design_day_columns <- raw_utility |> 
    filter(sheet == "DesignDay", 
           row == 8 & col %in% c(1:6, 8)) |> 
    select(character) |> 
    mutate(character = str_replace_all(character, "\\r\\n", " "),
           character = str_replace_all(character, "\\(Sum Columns 1-4\\)", ""), 
           character = str_replace(character, "CALCULATED", ""), 
           character = str_trim(character, side = "both")) |> 
    pull(character) |> 
    replace(1, "Year Label") |> 
    replace(2, "Year")
  
  design_day_values <- raw_utility |> 
    filter(sheet == "DesignDay", 
           row %in% c(9:15) & col %in% c(1:6, 8)) |> 
    select(row, col, content, character) |> 
    mutate(cell_value = case_when(
      col == 1 ~ character, 
      TRUE ~ content
    )) |> 
    select(row, col, cell_value) |> 
    pivot_wider(values_from = cell_value, names_from = col) |> 
    select(-c(row)) 
  
  colnames(design_day_values) <- design_day_columns
  
  design_day_clean <- design_day_values |> 
    mutate(Utility = utility_name) |> 
    mutate(across(-c(`Year Label`, "Year", "Utility"), as.numeric), 
           across(-c(`Year Label`, "Year", "Utility"), round)) |> 
    relocate("Utility", .before = `Year Label`) |> 
    mutate(utility_id = utility_id) |> 
    relocate(utility_id, .before = `Year Label`)
  
  return(design_day_clean)
  
}