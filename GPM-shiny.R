library(shiny)
library(dplyr)
library(readxl)
library(lubridate)
library(tidyr)
library(ggplot2)
library(forecast)

# Define forecasting function
forecasting_function <- function(data) {

  ts_data <- ts(data, frequency = 12) 
  ar <- arima(ts_data, order=c(2,1,1))
  forecast_data <- forecast(ar, h = 1)

  return(forecast_data)
}

cross_validation_function <- function(data, k = 6) {
  errors <- c()
  
  # Get the length of the time series data
  n <- length(data)
  
  # Loop through the last k months for cross-validation
  i <- k
  while (i > 0) {

    # Train on all data except the last i months
    train_data <- ts(data[1:(n - i)], frequency = 12)

    # Forecast the next month
    ar <- arima(train_data, order=c(2,1,1))
    forecasted <- forecast(ar, h = 1)

    # Calculate the error (e.g., Mean Absolute Percentage Error)
    actual <- data[n - i + 1]
    error <- accuracy(forecasted$mean, actual)[,5]
 
    # Append the error to the dataframe
    errors <- c(errors, error)
    i <- i - 1
    
  }
  
  data_date <- as.Date(data)
  
  errors_df <- data_frame(Time = data_date[(n-k+1):n],
                          Error = round(errors,2))

  return(errors_df)
}

# Define UI for the app
ui <- fluidPage(
  titlePanel("SPK: Forecast GPM"),
  
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Choose an Excel File", 
                accept = c(".xlsx")),
      actionButton("runForecast", "Run Forecast")
    ),
    
    mainPanel(
      h4(tags$b("Cross-Validation Errors")),
      dataTableOutput("cvTable"),
      fluidRow(
        column(width = 3, p(strong("Average Error:"))),
        column(width = 6, verbatimTextOutput("avgError", placeholder = 1))
      ),
      h4(tags$b("Forecast Plot")),
      plotOutput("forecastPlot"),
      h4(tags$b("Forecast Table")),
      tableOutput("forecastTable")
    )
  )
)

# Define server logic for the app
server <- function(input, output, session) {
  
  # Reactive expression to preprocess the data from all sheets
  data <- reactive({
    req(input$file)
    
    
    # Read all sheets into a list of dataframes
    df_list <- lapply(excel_sheets(input$file$datapath), function(x) {
      read_excel(input$file$datapath, sheet = x)
    })
    
    # Combine the dataframes from all sheets
    df_combined <- bind_rows(df_list, .id="Sheet") 
    df_combined$`Qty Sales (Kg)` <- as.numeric(df_combined$`Qty Sales (Kg)`)
    
    # Filter and drop NA
    df_filtered <- df_combined %>% mutate(Time = my(paste(Bulan, Tahun))) %>%
      filter(`Qty Sales (Kg)`>10)
    
    item_df <- df_filtered %>%
      group_by(Time) %>%
      summarise(mean_GPM = sum(`GPM (IDR)`))
    
    data <- item_df %>% drop_na()
    
    # Convert the data into a time series object
    ts_data <- ts(data$mean_GPM, 
                  start =  c(year(min(data$Time)), month(min(data$Time))), 
                  frequency = 12)
    
    return(ts_data)
  })
  
  # Observe the run forecast button click
  observeEvent(input$runForecast, {
    req(data())
    
    
    # Perform cross-validation on the last 6 months
    cv_errors <- cross_validation_function(data(), k = 6)

    # Output the cross-validation errors
    output$cvTable <- renderDataTable({
      cv_errors
    })
    
    output$avgError <- renderText({
      round(mean(cv_errors$Error),2)
    })
    
    # Run the forecasting function
    forecast_result <- forecasting_function(data())
    
    # Output the plot
    output$forecastPlot <- renderPlot({
      autoplot(forecast_result) +
        ggtitle("Forecast Results")
    })
    
    # Output the forecast table
    output$forecastTable <- renderTable({
      as.data.frame(forecast_result)
    })
  })
}

# Run the app
shinyApp(ui = ui, server = server)
