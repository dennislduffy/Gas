
utility_info <- function(directory, workbook){
  
  utility_sheet <- xlsx_cells(
    paste(directory, workbook, sep = "/"))
  
  utility_name <- utility_sheet |> 
    filter(sheet == "Registration" & address == "C9") |> 
    pull(character)
    
  utility_number <- utility_sheet |> 
      filter(sheet == "Registration" & address == "C5") |> 
      pull(content)
  
  report_year <- utility_sheet |> 
    filter(sheet == "Registration" & address == "C6") |> 
    pull(numeric)
  
  utility_info <- tibble(
    utility_name = utility_name,
    report_year = report_year,
    entity_id = utility_number
  )
  
  return(utility_info)
    
}








