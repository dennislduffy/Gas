county_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "GasByCounty")
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
  
  county_code <- raw_utility |> 
    filter(sheet == "GasByCounty",
           row %in% c(15:59) & col == 1 | row %in% c(15:56) & col == 5) |> 
    arrange(col, row) |> 
    select(content) |> 
    rename("county_code" = "content") 
  
  county_name <- raw_utility |> 
    filter(sheet == "GasByCounty", 
           row %in% c(15:59) & col == 2 | row %in% c(15:56) & col == 6) |>
    arrange(col, row) |> 
    select(character) |> 
    rename("county_name" = "character")
  
  mcf_delivered <- raw_utility |> 
    filter(sheet == "GasByCounty", 
           row %in% c(15:59) & col == 3 | row %in% c(15:56) & col == 7) |>
    arrange(col, row) |> 
    select(numeric) |> 
    rename("mcf_delivered" = "numeric")
  
  gasbycounty <- county_code |> 
    bind_cols(county_name) |> 
    bind_cols(mcf_delivered) |> 
    mutate(year = report_year, 
           utility = utility_name, 
           utility_id = utility_id) |> 
    filter(complete.cases(mcf_delivered))
  
  return(gasbycounty)
  
}