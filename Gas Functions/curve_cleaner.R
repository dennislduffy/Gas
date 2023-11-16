curve_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "LoadCurve")
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
  
  curve_names <- raw_utility |> 
    filter(sheet == "LoadCurve", 
           col %in% c(1:8, 11) & row == 19) |> 
    pull(character) |> 
    replace_na("Month") |> 
    str_replace_all("\\r\\n", " ") |> 
    str_replace_all(" \\(For The Entire Month\\)", "") |> 
    str_replace_all(" \\(MCF\\)", "") |> 
    str_trim(side = "both")
  
  curve_values <- raw_utility |> 
    filter(sheet == "LoadCurve", 
           col %in% c(1:8) & row %in% c(20:31)) |> 
    mutate(value = case_when(
      col == 1 ~ character, 
      TRUE ~ content
    )) |> 
    select(row, col, value) |> 
    pivot_wider(names_from = col, values_from = value) |> 
    mutate(across(.cols = -c(row, `1`), .fns = as.numeric)) |> 
    rowwise() |> 
    mutate(total = sum(across(3:9), na.rm = TRUE)) |> 
    select(-c(row))
  
  colnames(curve_values) <- curve_names
  
  load_curve <- curve_values |> 
    mutate(Utility = utility_name,
           utility_id = utility_id,
           Year = report_year) |> 
    relocate(Utility, .before = Month) |>
    relocate(utility_id, .before = Month) |> 
    relocate(Year, .before = Month)
  
  return(load_curve)
  
}