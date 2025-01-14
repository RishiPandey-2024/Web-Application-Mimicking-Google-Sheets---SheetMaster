library(shiny)
library(shinydashboard)
library(DT)
library(dplyr)
library(openxlsx)

# Define the UI
ui <- dashboardPage(
  dashboardHeader(title = "Google Sheets Clone", titleWidth = 300),
  dashboardSidebar(
    width = 250,
    sidebarMenu(
      menuItem("Spreadsheet", tabName = "spreadsheet", icon = icon("table")),
      menuItem("Formatting", tabName = "formatting", icon = icon("paint-brush")),
      menuItem("Save/Load", tabName = "saveload", icon = icon("save")),
      menuItem("Charts", tabName = "charts", icon = icon("chart-bar"))
    )
  ),
  dashboardBody(
    tabItems(
      # Spreadsheet Tab
      tabItem(
        tabName = "spreadsheet",
        fluidRow(
          box(
            width = 12, title = "Spreadsheet", status = "primary", solidHeader = TRUE,
            dataTableOutput("sheet")
          )
        ),
        fluidRow(
          column(4, textInput("formula", "Enter Formula (e.g., =SUM(A1:A3))")),
          column(4, actionButton("apply_formula", "Apply Formula", class = "btn-primary")),
          column(4, 
                 actionButton("add_row", "Add Row", class = "btn-success"), 
                 actionButton("add_col", "Add Column", class = "btn-success"))
        )
      ),
      
      # Formatting Tab
      tabItem(
        tabName = "formatting",
        fluidRow(
          box(
            title = "Formatting Options", width = 12, status = "info", solidHeader = TRUE,
            selectInput("font_style", "Font Style", choices = c("Normal", "Bold", "Italic")),
            selectInput("font_size", "Font Size", choices = c("12px", "14px", "16px", "18px")),
            actionButton("apply_formatting", "Apply Formatting", class = "btn-primary")
          )
        )
      ),
      
      # Save/Load Tab
      tabItem(
        tabName = "saveload",
        fluidRow(
          box(
            width = 6, title = "Load Spreadsheet", status = "warning", solidHeader = TRUE,
            fileInput("load_file", "Choose File", accept = c(".csv", ".rds", ".xlsx"))
          ),
          box(
            width = 6, title = "Save Spreadsheet", status = "success", solidHeader = TRUE,
            downloadButton("save_file", "Save", class = "btn-success")
          )
        )
      ),
      
      # Charts Tab
      tabItem(
        tabName = "charts",
        fluidRow(
          box(title = "Bar Chart", width = 6, status = "primary", solidHeader = TRUE, plotOutput("bar_chart")),
          box(title = "Line Chart", width = 6, status = "primary", solidHeader = TRUE, plotOutput("line_chart"))
        )
      )
    )
  )
)

# Define the server logic
server <- function(input, output, session) {
  # Reactive value to hold the spreadsheet data
  data <- reactiveVal(data.frame(matrix("", nrow = 10, ncol = 10, dimnames = list(NULL, LETTERS[1:10]))))
  
  # Render the spreadsheet
  output$sheet <- renderDataTable({
    datatable(data(), editable = TRUE, options = list(dom = "t", paging = FALSE))
  }, server = TRUE)
  
  # Handle cell edits
  observeEvent(input$sheet_cell_edit, {
    info <- input$sheet_cell_edit
    updated_data <- data()
    updated_data[info$row, info$col] <- info$value
    data(updated_data)
  })
  
  # Apply formulas
  observeEvent(input$apply_formula, {
    formula <- input$formula
    sheet_data <- data()
    
    if (grepl("^=SUM\\(", formula)) {
      range <- gsub("=SUM\\(|\\)", "", formula)
      cells <- strsplit(range, ":")[[1]]
      
      if (length(cells) == 2 && grepl("^[A-Z]\\d+$", cells[1]) && grepl("^[A-Z]\\d+$", cells[2])) {
        start_row <- as.numeric(gsub("\\D", "", cells[1]))
        end_row <- as.numeric(gsub("\\D", "", cells[2]))
        col <- gsub("\\d", "", cells[1])
        
        if (!is.na(start_row) && !is.na(end_row) && col %in% colnames(sheet_data)) {
          values <- as.numeric(sheet_data[start_row:end_row, col])
          sheet_data[1, 1] <- sum(values, na.rm = TRUE)
          data(sheet_data)
        } else {
          showNotification("Invalid range or out-of-bounds cells.", type = "error")
        }
      } else {
        showNotification("Invalid formula format. Use =SUM(A1:A3).", type = "error")
      }
    } else {
      showNotification("Unsupported formula. Only SUM is implemented.", type = "error")
    }
  })
  
  # Add Row and Column
  observeEvent(input$add_row, {
    updated_data <- data()
    updated_data <- rbind(updated_data, rep("", ncol(updated_data)))
    data(updated_data)
  })
  
  observeEvent(input$add_col, {
    updated_data <- data()
    new_col_name <- LETTERS[ncol(updated_data) + 1]
    updated_data[[new_col_name]] <- ""
    data(updated_data)
  })
  
  # Apply formatting
  observeEvent(input$apply_formatting, {
    font_style <- input$font_style
    font_size <- input$font_size
    showNotification(
      paste("Formatting applied:", font_style, "with font size", font_size),
      type = "message"
    )
  })
  
  # Save the spreadsheet
  output$save_file <- downloadHandler(
    filename = function() {"spreadsheet.xlsx"},
    content = function(file) {
      write.xlsx(data(), file, row.names = FALSE)
    }
  )
  
  # Load a spreadsheet
  observeEvent(input$load_file, {
    req(input$load_file)
    file_ext <- tools::file_ext(input$load_file$name)
    
    if (file_ext == "csv") {
      file_data <- read.csv(input$load_file$datapath, stringsAsFactors = FALSE)
      data(file_data)
    } else if (file_ext == "rds") {
      file_data <- readRDS(input$load_file$datapath)
      data(file_data)
    } else if (file_ext == "xlsx") {
      file_data <- read.xlsx(input$load_file$datapath)
      data(file_data)
    } else {
      showNotification("Unsupported file format. Use .csv, .rds, or .xlsx.", type = "error")
    }
  })
  
  # Generate bar chart
  output$bar_chart <- renderPlot({
    sheet_data <- data()
    # Convert the first row to numeric
    values <- suppressWarnings(as.numeric(sheet_data[1, ]))
    # Check if there are valid numeric values
    if (all(is.na(values))) {
      showNotification("No numeric data available for the bar chart.", type = "warning")
      return(NULL)
    }
    barplot(values, main = "Bar Chart", col = "blue")
  })
  
  # Generate line chart
  output$line_chart <- renderPlot({
    sheet_data <- data()
    # Convert the first row to numeric
    values <- suppressWarnings(as.numeric(sheet_data[1, ]))
    # Check if there are valid numeric values
    if (all(is.na(values))) {
      showNotification("No numeric data available for the line chart.", type = "warning")
      return(NULL)
    }
    plot(values, type = "l", main = "Line Chart", col = "red")
  })
}
# Run the application
shinyApp(ui = ui, server = server)
