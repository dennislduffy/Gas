pipeline_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "PipelineCo")
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
  
  pipeline <- raw_utility |> 
    filter(sheet == "PipelineCo", 
           col == 2) |> 
    select(character) |> 
    filter(complete.cases(character)) |> 
    rename(`Pipeline Company` = character) |> 
    mutate(Utility = utility_name,
           utility_id = utility_id,
           Year = report_year) |> 
    relocate(`Pipeline Company`, .after = Year)
  
  return(pipeline)
  
}
