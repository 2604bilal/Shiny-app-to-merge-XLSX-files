library(shiny)
library(readxl)
library(openxlsx)
library(dplyr)

ui <- fluidPage(
  titlePanel("XLSX Merger"),
  
  sidebarLayout(
    sidebarPanel(
      textInput("folderPath", "Folder Path:", placeholder = "Enter folder path"),
      textInput("sheetName", "Sheet Name:", value = "Raw Data", placeholder = "Enter sheet name"),  # Default to "Raw Data"
      actionButton("mergeButton", "Merge"),
      verbatimTextOutput("status") # To display messages
    ),
    mainPanel()
  )
)

server <- function(input, output) {
  
  observeEvent(input$mergeButton, {
    output$status <- renderPrint({"Processing..."}) # Initial message
    
    directory <- input$folderPath
    sheet_to_merge <- input$sheetName
    
    if (!dir.exists(directory)) {
      output$status <- renderPrint({paste("Error: Directory", directory, "not found.")})
      return() # Stop processing
    }
    
    xlsx_files <- list.files(directory, pattern = "\\.xlsx$", full.names = TRUE)
    
    if (length(xlsx_files) == 0) {
      output$status <- renderPrint({"Error: No XLSX files found in the directory."})
      return()
    }
    
    all_data <- list()
    merged_count <- 0
    missing_sheets <- character() # To store files with missing sheets
    
    for (file in xlsx_files) {
      tryCatch({
        sheets <- excel_sheets(file)
        
        if (sheet_to_merge %in% sheets) {
          data <- read_excel(file, sheet = sheet_to_merge, guess_max = 1000)  # Increase guess_max for better column type guessing
          
          # Convert all columns to character type to avoid type mismatch
          data <- data %>% mutate(across(everything(), as.character))
          
          if (nrow(data) > 0) {
            all_data <- append(all_data, list(data))
            merged_count <- merged_count + 1
          } else {
            print(paste("Warning: '", sheet_to_merge, "' sheet in", file, "is empty. Skipping."))
          }
          
        } else {
          missing_sheets <- c(missing_sheets, file) # Add to missing files list
          print(paste("Warning: '", sheet_to_merge, "' sheet not found in", file))
        }
      }, error = function(e) {
        print(paste("Error reading file", file, ":", e$message))
      })
    }
    
    if (length(missing_sheets) > 0) {
      output$status <- renderPrint({
        paste("Error: The following files are missing the '", sheet_to_merge, "' sheet:\n",
              paste(missing_sheets, collapse = "\n"), "\nProcess stopped.")
      })
      return() # Stop if any files are missing the sheet
    }
    
    if (length(all_data) > 0) {
      merged_data <- bind_rows(all_data)
      
      wb <- openxlsx::createWorkbook()
      addWorksheet(wb, "Merged Data")
      writeData(wb, sheet = "Merged Data", x = merged_data)
      
      output_file <- file.path(directory, "merged_data.xlsx")
      saveWorkbook(wb, output_file, overwrite = TRUE)
      
      output$status <- renderPrint({
        paste("Merged data saved to:", output_file, "\nTotal files merged:", merged_count)
      })
      
    } else {
      output$status <- renderPrint({
        paste("No non-empty '", sheet_to_merge, "' sheets found in any of the XLSX files.")
      })
    }
  })
}

shinyApp(ui = ui, server = server)
