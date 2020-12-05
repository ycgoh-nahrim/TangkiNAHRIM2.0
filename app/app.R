library(shiny)
library(tidyverse)
library(lubridate)
library(plotly)
library(leaflet)
library(shinydashboard)
library(htmlwidgets)
library(htmltools)
library(mapview)
library(shinyBS)
#library(shinycssloaders)
#library(shinyjs)
library(shinybusy)
library(markdown)
library(openxlsx)


#get rainfall data
rain_db <- read.csv("rainfall_msia2.csv", 
                    stringsAsFactors = FALSE, header = TRUE, sep=",")

#get rainfall station data
rainfall_stn <- read.csv("rain_stn_msia2.csv", 
                         stringsAsFactors=F)

rainfall_stn <- rainfall_stn %>%
  mutate(popup = str_c(STATION_NO,
                       STATION_NA,
                       sep = "<br/>"))

#go to top of page function
#jscode <- "shinyjs.toTop = function() {document.body.scrollTop = 0;}"


##UI#############
ui <- fluidPage(
  
  tags$head(includeHTML(("google-analytics.html"))),
  
  #force desktop view in mobile screen
  tags$head(tags$meta(name = "viewport", 
                      content = "width=1600"), 
            uiOutput("body")),
  
  
  #add busy indicator
  add_busy_bar(color = "#0048FF"),
  
  
  titlePanel(title=div(img(src="Nahrim.png"), "Tangki NAHRIM 2.0"),
             windowTitle = "Tangki NAHRIM 2.0"),
  
  p("Simple web app to estimate the most optimal rainwater harvesting tank size",
    style = "font-family: 'Roboto Light'; color: gray27"),

  
  sidebarPanel(
    
    
    
    wellPanel(
      
      h4("Rainfall at Location"),
      
      #station number
      #uiOutput("stn_Output"),
      
      #selectInput(inputId = "stn_Output",
      #            label = "Select station nearest to your property from the map (click 'Locate Me' button, then select red marker)",
      #            choices = sort(unique(rain_db$stn_no)),
      #            selected = " ",
      #            multiple = F), 
      selectizeInput(inputId = "stn_Output",
                     #label = "Select station nearest to your property from the map (click 'Locate Me' button, then select red marker)",
                     label = tags$div(HTML('Select station nearest to your property from the map (click <i class="fa fa-crosshairs""></i> to zoom in to your location, then select red marker)')),
                     choices = sort(unique(rain_db$stn_no)),
                     multiple = F,
                     options = list(maxItems = 1, placeholder = 'Enter Station Number (if known)',
                                    onInitialize = I('function() { this.setValue(""); }'))),
      #bsTooltip("stn_Output", 
      #          "Station Number. Click Locate Me button on the map to zoom into your current location"),
      
      
      #load custom rainfall data
      p("If you have your own daily rainfall data (csv only)"),
      
      #enter location/station name text input?
      
      fileInput('target_upload', 'Choose CSV file',
                accept = c(
                  'text/csv',
                  'text/comma-separated-values',
                  '.csv')
                ),
      #checkboxInput("header", "Header", TRUE),
      
      # reduce margin between file input and note
      div(style = "margin-top: -30px"),
      
      p("Note: First column is 'Date' ('dd/mm/yyyy' format), second column is 'Depth' (rainfall value in mm)",
        style = "font-family: 'Calibri Light'; color: gray27; font-size: 1.5rem"),
      
      #radioButtons("separator","Separator: ",choices = c(";",",",":"), selected=";",inline=TRUE),
      #DT::dataTableOutput("sample_table")
      
      #read rainfall data
      actionButton("Read_action", 
                   label = "Click Here to Proceed", class = "btn-primary")
                   #style = "background-color: #004C99; color: #FFFFFF")

    ),
    
    #hr(),
    wellPanel(
      h4("Roof Information"),
      
      fluidRow(
        column(6,
               
               #roof length (m)
               numericInput("roof_length_input", "Roof Length (m)", 10)
               
               ),
        
        column(6,
               
               #runoff coefficient
               numericInput("runoff_coef_input", "Roof Runoff Coefficient", 0.8,
                            min = 0.5, max = 1, step = 0.05), 
               bsTooltip("runoff_coef_input", "0.8 for concrete roof, 0.9 for zink/metal roof, or other value less than 1")
              
               )
      ),
      
      fluidRow(
        column(6,
               
               #roof width (m)
               numericInput("roof_width_input", "Roof Width (m)", 10)
               
        ),
        
        column(6,
               
               #first flush (mm)
               numericInput("first_flush_input", "First Flush (mm)", 1),
               bsTooltip("first_flush_input", "1 mm is recommended to wash away dirt from roof at the beginning of rain")
               
        )
          
      )

    ),
    
    wellPanel(
      
      h4("Water Demand"),
      
      #water demand (l)
      numericInput("water_demand_input", "Potential amount of harvested rainwater to be used or total Water Demand (litres per day)", 150),
      bsTooltip("water_demand_input", "For toilet flushing, gardening, car washing, general cleaning etc")
     
      ),
      

    
    wellPanel(
      h4("Tank Capacity"),
      
      fluidRow(
        column(6,
               
               #tank capacity - min (m3)
               numericInput("tank_cap_min_input", 
                            HTML(paste("Smallest size considered ( m", tags$sup(3), " )", 
                                       sep = "")), 
                            1),
               bsTooltip("tank_cap_min_input", "Minimum tank capacity"),
               
               
               #tank capacity - interval (m3)
               numericInput("tank_cap_interval_input", 
                            HTML(paste("Size in-between", tags$br(), "( m", tags$sup(3), " )", 
                                       sep = "")), 
                            1),
               bsTooltip("tank_cap_interval_input", "Tank capacity interval")

               ),
        
        column(6,
               
               #tank capacity - max (m3)
               numericInput("tank_cap_max_input", 
                            HTML(paste("Largest size considered ( m", tags$sup(3), " )", 
                                       sep = "")), 
                            10),
               bsTooltip("tank_cap_max_input", "Maximum tank capacity")
               
               )
      )

    ),
    
    
    
    #calculate
    actionButton("Calculate_action", 
                 label = "Calculate",
                 class = "btn-primary"),
    
    br(),

    p("Note: If chart looks distorted, refresh the page and try again.",
      style = "font-family: 'Calibri Light'; color: gray27; font-size: 1.5rem"),
    
    hr(),
    
    p("Results for advanced analysis can be downloaded here (in Excel format)",
      style = "color: gray27; font-size: 1.5rem"),
    
    
    #Download results
    downloadButton("downloadResults", "Download Results")
    

  ),
  
  mainPanel(tabsetPanel(
    
    id="inTabset",
    
    tabPanel(title = "Map", 
             leafletOutput("mymap",
                           height = 600)
    ),
              
    tabPanel(title = "Rainfall", 
             h4("Total Annual Rainfall"),
             #withSpinner(plotlyOutput("rainfall_annual_Plot")),
             plotlyOutput("rainfall_annual_Plot"),
             h4("Monthly Rainfall"),
             #withSpinner(plotlyOutput("rainfall_mth_Plot")),
             plotlyOutput("rainfall_mth_Plot")
             #tableOutput("rainfall_annual_Table")

             #fluidRow(
             #  column(12, tableOutput("rainfall_annual_Table")))
             ),
              
    tabPanel(title = "Results", 
             h4("Efficiency vs. Tank Sizes"),
             p("Please hover over each point to check the benefit (efficiency) of the each tank size",
               style = "font-family: 'Roboto Light'; color: gray27"),
             #withSpinner(plotlyOutput("WS_Sto_eff_Plot")),
             plotlyOutput("WS_Sto_eff_Plot"),
             
             #textOutput("results_text"),
             
             #h4("Results: Efficiency"),
             #tableOutput("WS_Sto_eff_Table"),
             
             h4("Tank Volume Percentage Time"),                        
             plotlyOutput("tank_volume_plot")
             
             ),
    
    tabPanel(title = "FAQ", 
             
             #tabsetPanel(
               
               #tabPanel("En",
                 
                 includeHTML("FAQ.html")
                 
               #),
               
               #tabPanel("BM",
                 
                 #includeHTML("FAQ_BM.html")
                 
               #)
               
             #)
             
             
             
             ),
    
    tabPanel(title = "About", 
             
             #tabsetPanel(
               
               #tabPanel("En",
                        
                        includeHTML("About.html")
                        
               #),
               
               #tabPanel("BM",
                        
                        #includeHTML("About_BM.html")
                        
              # )
               
             #)
             
             )
    
    )
    
    )
  
)


##SERVER#############
server <- function(input, output, session) {
  
  #GENERAL INFROMATION
  gen_info <- reactiveValues()
  observe({
    
    # app version
    gen_info$version_no <- "Tangki NAHRIM 2.0"
    gen_info$notes <- paste0("Generated by: ", gen_info$version_no)
    
  })
  
  
  #MAP
  output$mymap <- renderLeaflet({
    leaflet(rainfall_stn) %>% 
      addTiles() %>%
      addCircleMarkers(lng = ~X, lat = ~Y, 
                       layerId = ~STATION_NO, 
                       label=~STATION_NA,
                       color = "red", 
                       radius = 3,
                       weight = 7,
                       popup = ~popup) %>%
      addEasyButton(easyButton(
        icon = "fa-globe", title = "Zoom out",
        onClick = JS("function(btn, map){ map.setZoom(7);}"))) %>%
      addEasyButton(easyButton(
        icon = "fa-crosshairs", title = "Locate Me",
        onClick = JS("function(btn, map){ map.locate({setView: true, maxZoom: 12}); }")))
    
  })
  
  #MAP UPDATE SELECTION (SELECT STATION)
  
  ##from select input
  ##selected station data
  filtered <- reactive({
    rainfall_stn[rainfall_stn$STATION_NO == input$stn_Output,]
  })
  
  ###update map according to station number
  mymap_proxy <- leafletProxy("mymap")
  
  observe({
    fdata <- filtered()
    mymap_proxy %>%
      #clearMarkers() %>%
      #addMarkers(lng = fdata$X, lat = fdata$Y, label = fdata$STATION_NA) %>%
      #addLabelOnlyMarkers(lng = fdata$X, lat = fdata$Y) %>%
      addPopups(lng = fdata$X, lat = fdata$Y, 
                popup = fdata$popup,
                options = popupOptions(closeButton = FALSE)) %>%
      flyTo(lng = fdata$X, lat = fdata$Y, zoom = 12)
  })
  

  
  ##from select marker
  observeEvent(input$mymap_marker_click, {
    
    click <- input$mymap_marker_click
    station <- rainfall_stn[which(rainfall_stn$Y == click$lat & 
                                    rainfall_stn$X == click$lng), ]$STATION_NO
    
    #update input according to selected marker
    updateSelectInput(session = session, 
                      inputId = "stn_Output",
                      choices = sort(unique(rain_db$stn_no)), 
                      selected = c(input$STATION_NO, station))
    
  })
  
  
  
  #DAILY RAINFALL DATA SELECTION
  
  raindata_sel_df <- reactive({

    inFile <- input$target_upload
    
    if (is.null(inFile)) {
      
      #from csv database
      raindata_df <- rain_db %>%
        filter(stn_no == input$stn_Output) %>% 
        select(Date, Depth) %>%
        arrange(Date)
      raindata_df$Date <- as.Date(raindata_df$Date, format = "%Y-%m-%d")
      raindata_df
      

    } else {
      
      #from user input
      raindata_df <- read.csv(inFile$datapath)
                              #header = input$header)
      colnames(raindata_df)[1:2] <- c("Date", "Depth")
      raindata_df$Date <- as.Date(raindata_df$Date, format = "%d/%m/%Y")
      raindata_df

    }

  })
  
  
  #READ AND DISPLAY RAINFALL DATA
  
  #rainfall summary for annual and monthly in dataframe
  rain_summary <- reactive({
    
    #pass reactive variable to dataframe (selected daily data df)
    raindata_sel <- raindata_sel_df()
    
    
    #annual rainfall
    raindata_yr <- raindata_sel %>%
      mutate(year = year(Date)) %>%
      group_by(year) %>%
      summarise(sum_precip = sum(Depth))
    raindata_yr
    
    #annual rainfall for excel output
    raindata_yr2 <- raindata_yr %>%
      rename(Year = year,
             Rainfall_mm = sum_precip)
    raindata_yr2$Rainfall_mm <- lapply(raindata_yr2$Rainfall_mm, 
                                            round, digits = 1)
    raindata_yr2$Rainfall_mm <- as.numeric(raindata_yr2$Rainfall_mm)
    raindata_yr2

    
    #monthly rainfall
    raindata_mth <- raindata_sel %>%
      mutate(year= year(Date), month = month(Date)) %>%
      group_by(year, month) %>%
      summarise(sum_precip = sum(Depth)) %>%
      group_by(month) %>%
      summarise(mth_precip = mean(sum_precip))
    raindata_mth
    
    #monthly rainfall for excel output
    raindata_mth2 <- raindata_mth %>%
      rename(Month = month,
             Rainfall_mm = mth_precip)
    raindata_mth2$Rainfall_mm <- lapply(raindata_mth2$Rainfall_mm, 
                                             round, digits = 1)
    raindata_mth2$Rainfall_mm <- as.numeric(raindata_mth2$Rainfall_mm)
    raindata_mth2
    
    
    #list
    list(rain_yr_chart = raindata_yr,
         rain_yr_table = raindata_yr2,
         rain_mth_chart = raindata_mth,
         rain_mth_table = raindata_mth2)
 
    
  })
  
  
  #rainfall summary values
  rain_summary_val <- reactiveValues()
  
  observe({
    
    #annual rainfall df
    raindata <- rain_summary()[['rain_yr_chart']]
    #daily data df
    raindata_sel <- raindata_sel_df()
    
    #calculation/summary
    raindata$year <- as.numeric(raindata$year)
    rain_summary_val$max_year <- max(raindata$year, na.rm = T)
    rain_summary_val$min_year <- min(raindata$year, na.rm = T)
    rain_summary_val$angle_precip <- ifelse(rain_summary_val$max_year - rain_summary_val$min_year + 1 > 20, 
                                            90, 
                                            0)
    rain_summary_val$total_yr <- rain_summary_val$max_year - rain_summary_val$min_year + 1
    rain_summary_val$yr_w_data <- length(unique(raindata$year))
    rain_summary_val$day_w_data <- length(unique(raindata_sel$Date))
    
  })
  
  
  
  
  #annual rainfall chart
  rain_yr_plot <- function(){
    
    #pass reactive variable to dataframe
    #annual rainfall df
    raindata <- rain_summary()[['rain_yr_chart']]
    
    r_yr <-
    
    ggplotly(
      ggplot(raindata, aes(x = year, y = sum_precip)) +
        geom_bar(aes(text = paste0('Total rainfall in year <b>',
                                   year,
                                   '</b> is <b>',
                                   sprintf("%0.1f", sum_precip),
                                   'mm</b>')),
                 stat = "identity", fill = "steelblue") +
        theme_bw(base_size = 10) +
        scale_x_continuous(name= "Year", 
                           breaks = seq(rain_summary_val$min_year, 
                                        rain_summary_val$max_year, 
                                        by = 1), 
                           minor_breaks = NULL) + #x axis format
        scale_y_continuous(name= "Annual Rainfall (mm)",
                           #breaks = seq(0, 3500, by = 500), 
                           minor_breaks = NULL) + #y axis format
        geom_hline(aes(yintercept = mean(sum_precip)), 
                   color="black", 
                   alpha=0.3, 
                   size=1) + #avg line
        geom_text(aes(rain_summary_val$min_year+1,
                      mean(sum_precip),
                      label = paste("Average = ", sprintf("%0.0f", mean(sum_precip)), "mm"), 
                      vjust = -0.5, hjust = 0), 
                  size=3.5, 
                  #family="Roboto Light", 
                  fontface = 1, 
                  color="grey20") + #avg line label
        theme(text=element_text(color="grey20"),
              panel.grid.major.x = element_blank(),
              axis.text.x = element_text(angle = rain_summary_val$angle_precip, 
                                         hjust = 0.5))
      ,
      tooltip = "text")
    
    r_yr
    

  }

  
  
  #monthly rainfall chart
  rain_mth_plot <- function(){
    
    #pass reactive variable to dataframe
    #monthly rainfall df
    raindata <- rain_summary()[['rain_mth_chart']]
    
    r_mth <-
      
      ggplotly(
        ggplot(raindata, aes(x = month, y = mth_precip)) +
          geom_bar(aes(text = paste0('Long-term average rainfall in <b>',
                                     month.abb,
                                     '</b> is <b>',
                                     sprintf("%0.1f", mth_precip),
                                     'mm</b>')),
                   stat = "identity", fill = "steelblue") +
          theme_bw(base_size = 10) +
          scale_x_continuous(name = "Month",
                             breaks = seq(1, 12, by = 1), 
                             minor_breaks = NULL, 
                             labels=month.abb) + #x axis format
          scale_y_continuous(name = "Monthly Rainfall (mm)",
                             #breaks = seq(0, 350, by = 50), 
                             minor_breaks = NULL) + #y axis format
          theme(text=element_text(color="grey20"),
                panel.grid.major.x = element_blank())
        ,
        tooltip = "text")
    
    r_mth
    
  }
    
    
  #BUTTON - read data
  #output 
  output$Read_action <- renderPrint({ input$Read_action })
  
  observeEvent(input$Read_action,{
    
    #annual rainfall bar chart
    output$rainfall_annual_Plot <- renderPlotly({
      
      rain_yr_plot()
      
    })
    
    
    #monthly rainfall bar chart
    output$rainfall_mth_Plot <- renderPlotly({
      
      rain_mth_plot()
      
    })
    
    
    #annual rainfall table
    output$rainfall_annual_Table <- renderTable({
      
      #pass reactive variable to dataframe
      raindata_sel <- rain_summary()[['rain_yr_chart']]
      
      
      raindata <- raindata_sel %>%
        rename('Rainfall (mm)' = sum_precip,
               Year = year)
      
      raindata
      
    }, 
    
    #table properties
    rownames = TRUE,
    hover = TRUE,
    align = '?', 
    digits = 0)
    
    
    #switch to tab
    updateTabsetPanel(session, "inTabset",selected = "Rainfall")
    #js$toTop();
    
    
  })
  

  
  
  #TANGKI NAHRIM CALCULATION
  
  tangki_nahrim <- reactive({
    
    #pass reactive variable to dataframe
    raindata_sel <- raindata_sel_df()
    
    
    #parameters
    station_no <- as.numeric(input$stn_Output) #station number
    roof_length <- as.numeric(input$roof_length_input) #in meters
    roof_width <- as.numeric(input$roof_width_input) #in meters
    Runoff_coef <- as.numeric(input$runoff_coef_input) #usually 0.8 for concrete tiles, 0.9 for zink/metals
    First_flush <- as.numeric(input$first_flush_input) #in mm
    Water_demand_l <- as.numeric(input$water_demand_input) #in litres
    Tank_capacity_1 <- as.numeric(input$tank_cap_min_input) #in m3, first value of tank capacity
    Tank_capacity_10 <- as.numeric(input$tank_cap_max_input) #in m3, last value of tank capacity
    Tank_capacity_interval <- as.numeric(input$tank_cap_interval_input) #in m3, tank capacity interval
    
    raindata <- raindata_sel 
    
    
    #CALCULATION#####################
    #tank capacity iteration
    tank_mul <- (Tank_capacity_10 - Tank_capacity_1)/Tank_capacity_interval + 1
    
    #roof area
    raindata <- raindata %>%
      mutate(Roof_area = roof_length*roof_width)
    
    #runoff
    #Khastagir, A. and Jayasuriya, N., 2010. 
    #Optimal sizing of rain water tanks for domestic water conservation. 
    #Journal of Hydrology, 381 (3), 181-188.
    raindata <- raindata %>%
      mutate(Runoff = pmax(((raindata$Depth - First_flush)/1000*Roof_area*Runoff_coef),
                           0))
    #Water_demand
    raindata <- raindata %>%
      mutate(Water_demand = Water_demand_l/1000)
    
    
    #####################
    #initialization for dataframe calculation
    raindata$Tank_capacity <- Tank_capacity_1 #first tank capacity
    raindata$Active_storage <- 0 #tank starts empty
    raindata$Yield <- 0
    raindata$yield1 <- 0
    raindata$Spillage <- 0
    raindata$active_storage1 <- 0
    raindata$active_storage2 <- 0
    raindata$spillage1 <- 0
    
    #####################
    #iterate for range of Tank_capacity
    #ACTIVE STORAGE & YIELD (YAS)
    #Jenkins, D., Pearson, F., Moore, E., Kim, S.J., and Valentine, R., 1978. 
    #Feasibility of Rainwater Collection Systems in California. Contribution No 173, 
    #Californian Water Resources Centre, University of California.
    #SPILLAGE
    #Campisano, A. and Modica, C., 2015. 
    #Appropriate resolution timescale to evaluate water saving and retention potential of 
    #rainwater harvesting for toilet flushing in single houses. Journal of Hydroinformatics, 17 (3), 331-346.
    
    datalist = list() #for combination
    counter <- 0
    for(j in seq(Tank_capacity_1, Tank_capacity_10, Tank_capacity_interval)) {
      raindata$Tank_capacity <- j
      for (i in 2:nrow(raindata)) {
        raindata$yield1[i] = (raindata$Active_storage[i-1] + raindata$Runoff[i])
        raindata$Yield[i] = min(c(raindata$yield1[i],raindata$Water_demand[i]))
        raindata$active_storage1[i] = (raindata$Active_storage[i-1] + raindata$Runoff[i] - raindata$Yield[i])
        raindata$active_storage2[i] = raindata$Tank_capacity[i] - raindata$Yield[i]
        raindata$Active_storage[i] = min(c(raindata$active_storage1[i],raindata$active_storage2[i]))
        raindata$spillage1[i] = (raindata$Active_storage[i-1] + raindata$Runoff[i] - raindata$Tank_capacity[i])
        raindata$Spillage[i] = max(c(raindata$spillage1[i], 0))
      }
      counter <- counter + 1
      datalist[[counter]] <- raindata
    }
    #combine all iterations of Tank_capacity
    tankdata = do.call(rbind, datalist)
    
    #################################
    
    #Yield_demand_d
    tankdata <- tankdata %>%
      mutate(Yield_demand_d = ifelse(Yield == Water_demand, 1, 0))
    
    #tank volume percentage
    tankdata <- tankdata %>%
      mutate(Tank_100 = ifelse(Active_storage/Tank_capacity >= 0.75, 1, 0),
             Tank_75 = ifelse(Active_storage/Tank_capacity >= 0.5, 
                              (ifelse(Active_storage/Tank_capacity < 0.75, 1, 0)), 0),
             Tank_50 = ifelse(Active_storage/Tank_capacity >= 0.25, 
                              (ifelse(Active_storage/Tank_capacity < 0.5, 1, 0)), 0),
             Tank_25= ifelse(Active_storage/Tank_capacity > 0, 
                             (ifelse(Active_storage/Tank_capacity < 0.25, 1, 0)), 0),
             Tank_0= ifelse(Active_storage/Tank_capacity == 0, 1, 0))
    
    
    
    
    
    ##################
    # add a year column to data.frame
    tankdata <- tankdata %>%
      mutate(year = year(Date))
    # add a month column to data.frame
    tankdata <- tankdata %>%
      mutate(month = month(Date))
    

    #CALCULATION FOR ANALYSIS##############################################
    
    # calculate the sum rain days and no rain days for each year
    raind_yr <- tankdata %>%
      group_by(year) %>%
      summarise(sum_rain = sum(Runoff > 0)/tank_mul, 
                sum_norain = sum(Runoff == 0)/tank_mul) #tank sizes iteration
    #pivot for stacked bar plot
    raind_yr2 <- raind_yr %>%
      gather(type, no_day, sum_rain:sum_norain)
    #cleanup for output
    raind_yr3 <- raind_yr %>%
      rename(Year = year,
             "Rain" = sum_rain,
             "No Rain" = sum_norain)
    
    
    
    #calculate rain days and no rain days for each month, each year
    raind_mth_yr <- tankdata %>%
      group_by(month, year) %>%
      summarise(sum_rain = sum(Runoff > 0)/tank_mul, 
                sum_norain = sum(Runoff == 0)/tank_mul) #tank sizes iteration
    raind_mth <- raind_mth_yr %>%
      group_by(month) %>%
      summarise(mth_rain = mean(sum_rain), mth_norain = mean(sum_norain))
    #pivot for stacked bar plot
    raind_mth2 <- raind_mth %>%
      gather(type, no_day, mth_rain:mth_norain)
    #cleanup for output
    raind_mth3 <- raind_mth %>%
      rename(Month = month,
             "Rain" = mth_rain,
             "No Rain" = mth_norain)
    raind_mth3$"Rain" <- lapply(raind_mth3$"Rain", round, digits = 1)
    raind_mth3$"Rain" <- as.numeric(raind_mth3$"Rain")
    raind_mth3$"No Rain" <- lapply(raind_mth3$"No Rain", round, digits = 1)
    raind_mth3$"No Rain" <- as.numeric(raind_mth3$"No Rain")
    
    
    
    #calculate percentage no rain days
    day_norain <- tankdata %>%
      summarise(sum_norain = sum(Runoff == 0)/tank_mul)
    
    
    
    #calculate yield and spillage volume for each tank size
    vol_yield_spill_yr <- tankdata %>%
      group_by(Tank_capacity, year) %>%
      summarise(sum_yield = sum(Yield), sum_spill = sum(Spillage))
    vol_yield_spill <- vol_yield_spill_yr %>%
      group_by(Tank_capacity) %>%
      summarise(yield_yr = mean(sum_yield), spill_yr = mean(sum_spill))
    #pivot for stacked bar plot
    vol_yield_spill2 <- vol_yield_spill %>%
      gather(YS, volume, yield_yr:spill_yr)
    #cleanup for output
    vol_yield_spill3 <- vol_yield_spill %>%
      rename("Yield (m3/yr)" = yield_yr,
             "Spill (m3/yr)" = spill_yr)
    vol_yield_spill3$"Yield (m3/yr)" <- lapply(vol_yield_spill3$"Yield (m3/yr)", 
                                               round, digits = 1)
    vol_yield_spill3$"Spill (m3/yr)" <- lapply(vol_yield_spill3$"Spill (m3/yr)", 
                                               round, digits = 1)
    vol_yield_spill3$"Yield (m3/yr)" <- as.numeric(vol_yield_spill3$"Yield (m3/yr)")
    vol_yield_spill3$"Spill (m3/yr)" <- as.numeric(vol_yield_spill3$"Spill (m3/yr)")
    
    
    
    #calculate yield number of days for each tank size
    day_yield_yr <- tankdata %>%
      group_by(Tank_capacity, year) %>%
      summarise(sum_yield = sum(Yield_demand_d))
    day_yield <- day_yield_yr %>%
      group_by(Tank_capacity) %>%
      summarise(yield_yr = mean(sum_yield))
    
    #calculate spillage number of days for each tank size
    day_spill_yr <- tankdata %>%
      group_by(Tank_capacity, year) %>%
      summarise(count_spill = n(), sum_spill = sum(Spillage > 0))
    day_spill <- day_spill_yr %>%
      group_by(Tank_capacity) %>%
      summarise(spill_yr = mean(sum_spill))
    
    #combine yield and spillage day for each tank size
    day_yield_spill <- full_join(day_yield, day_spill, by = "Tank_capacity")
    #pivot
    day_yield_spill2 <- day_yield_spill %>%
      gather(YS, No_day, yield_yr:spill_yr)
    #cleanup for output
    day_yield_spill3 <- day_yield_spill %>%
      rename("Yield (day/yr)" = yield_yr,
             "Spill (day/yr)" = spill_yr)
    day_yield_spill3$"Yield (day/yr)" <- lapply(day_yield_spill3$"Yield (day/yr)", 
                                               round, digits = 1)
    day_yield_spill3$"Spill (day/yr)" <- lapply(day_yield_spill3$"Spill (day/yr)", 
                                               round, digits = 1)
    day_yield_spill3$"Yield (day/yr)" <- as.numeric(day_yield_spill3$"Yield (day/yr)")
    day_yield_spill3$"Spill (day/yr)" <- as.numeric(day_yield_spill3$"Spill (day/yr)")

    
    
    #calculate water-saving efficiency
    WS_eff_tank <- tankdata %>%
      group_by(Tank_capacity) %>%
      summarise(WS_eff = sum(Yield)/sum(Water_demand)*100)
    
    #calculate storage efficiency
    Sto_eff_tank <- tankdata %>%
      group_by(Tank_capacity) %>%
      summarise(Sto_eff = (1 - (sum(Spillage)/sum(Runoff)))*100)
    
    #combine both efficiency measures
    WS_Sto_eff <- full_join(WS_eff_tank, Sto_eff_tank, by = "Tank_capacity")
    #pivot
    WS_Sto_eff2 <- WS_Sto_eff %>%
      gather(eff, value, WS_eff:Sto_eff)
    #join yield table
    WS_Sto_eff2 <- full_join(WS_Sto_eff2, vol_yield_spill3, by = "Tank_capacity")
    #add columns for hover text
    WS_Sto_eff2 <- WS_Sto_eff2 %>%
      mutate(hovertext = ifelse(WS_Sto_eff2$eff == 'WS_eff',
                                paste0('<b>',
                                       sprintf("%0.1f", value),
                                       '% of water demand</b> can be met for tank size <b>',
                                       Tank_capacity,
                                       ' m<sup>3</sup></b><br>',
                                       '(<b>', 
                                       WS_Sto_eff2$'Yield (m3/yr)',
                                       ' m<sup>3</sup> of water</b> can be saved per year)'),
                                paste0('<b>',
                                       sprintf("%0.1f", value),
                                       '%</b> of rainwater from roof can be used for tank size <b>',
                                       Tank_capacity,
                                       ' m<sup>3</sup></b><br>')))
    #replace
    WS_Sto_eff2$eff <- WS_Sto_eff2$eff %>% 
      str_replace("WS_eff", "Water-Saving Efficiency") %>%
      str_replace("Sto_eff", "Storage Efficiency")
    
    
    
    #format number
    WS_Sto_eff4 <- WS_Sto_eff
    WS_Sto_eff4$WS_eff <- lapply(WS_Sto_eff4$WS_eff, round, digits = 1)
    WS_Sto_eff4$Sto_eff <- lapply(WS_Sto_eff4$Sto_eff, round, digits = 1)
    #table for output
    WS_Sto_eff4 <- WS_Sto_eff4 %>%
      rename("Water-Saving Efficiency" = WS_eff) %>%
      rename("Storage Efficiency" = Sto_eff)
    WS_Sto_eff4$"Water-Saving Efficiency" <- as.numeric(WS_Sto_eff4$"Water-Saving Efficiency")
    WS_Sto_eff4$"Storage Efficiency" <- as.numeric(WS_Sto_eff4$"Storage Efficiency")
    
    
    
    
    #calculate percentage tank volume
    tank_vol <- tankdata %>%
      group_by(Tank_capacity) %>%
      summarise(tankvol_100 = sum(Tank_100)/n()*100,
                tankvol_75 = sum(Tank_75)/n()*100,
                tankvol_50 = sum(Tank_50)/n()*100,
                tankvol_25 = sum(Tank_25)/n()*100,
                tankvol_0 = sum(Tank_0)/n()*100
                )
    #pivot for stacked bar plot
    tank_vol2 <- tank_vol %>%
      gather(vol, perc, tankvol_0:tankvol_100) %>%
      mutate(label_vol = vol) 
    tank_vol2$label_vol <- tank_vol2$label_vol %>% 
      str_replace("tankvol_100", "75% - 100%") %>%
      str_replace("tankvol_75", "50% - 75%") %>% 
      str_replace("tankvol_50", "25% - 50%") %>%
      str_replace("tankvol_25", "0% - 25%")  %>%
      str_replace("tankvol_0", "0 (Empty)")
      
    #change factor level to reorder tank vol
    tank_vol_order <- c("75% - 100%", "50% - 75%", 
                        "25% - 50%", "0% - 25%", "0 (Empty)")
    tank_vol2$label_vol <- factor(tank_vol2$label_vol, 
                                  levels = tank_vol_order)
    tank_vol2 <- tank_vol2[order(tank_vol2$label_vol),]
    
    #add tooltips
    tank_vol2 <- tank_vol2 %>%
      mutate(hovertext = paste0("Tank volume at <b>",
                                label_vol,
                                "</b><br><b>",
                                sprintf("%0.1f", perc),
                                "%</b> of the time for tank size <b>",
                                Tank_capacity,
                                " m<sup>3</sup></b>")
             )
    tank_vol2$label_vol <- as.factor(tank_vol2$label_vol)
    
    
    #cleanup for output in Excel
    tank_vol3 <- tank_vol2
    tank_vol3$perc <- lapply(tank_vol3$perc, round, digits = 1)
    tank_vol3$perc <- as.numeric(tank_vol3$perc)
    
    #output for Excel 
    tank_vol4 <- tank_vol3 %>%
      select("Tank_capacity", "label_vol", "perc") %>%
      spread(key = "label_vol",
             value = "perc")
    #reorder columns
    tank_vol4 <- tank_vol4[, c("Tank_capacity", "0 (Empty)", "0% - 25%",
                               "25% - 50%", "50% - 75%", "75% - 100%")]
    
    
    
    #LIST OF DATAFRAMES
    list(
      #rain days
      raind_yr_chart = raind_yr2,
      raind_yr_table = raind_yr3,
      raind_mth_chart = raind_mth2,
      raind_mth_table = raind_mth3,
      
      #yield-spill by volume
      vol_yield_spill_chart = vol_yield_spill2,
      vol_yield_spill_table = vol_yield_spill3,
      
      #yield-spill by day
      day_yield_spill_chart = day_yield_spill2,
      day_yield_spill_table = day_yield_spill3,
      
      #efficiencies
      WS_Sto_eff_chart = WS_Sto_eff2, #for chart
      WS_Sto_eff_table = WS_Sto_eff4, #for table
      
      #tank volume percentage
      tank_vol_chart = tank_vol2,
      tank_vol_table = tank_vol4
      
    )
    
    
  })
  
  
  #CHARTS
  
  
  ##Water-saving and storage efficiency
  ws_sto_plot <- function(){
    
    
    #efficiency data frame
    WS_Sto_eff2 <- tangki_nahrim()[['WS_Sto_eff_chart']]
    
    WS_Sto_chart <-
      
      ggplotly(
        
        ggplot(WS_Sto_eff2, aes(x = Tank_capacity, y = value, group = eff)) +
          geom_line(aes(color = eff), size = 1) +
          geom_point(aes(color = eff,
                         text = hovertext), 
                     size = 2) +
          scale_color_manual(values = c("slategrey", "turquoise3")) +
          theme_bw(base_size = 10) +
          scale_x_continuous(name = "Tank Capacity (m<sup>3</sup>)", 
                             breaks = seq(input$tank_cap_min_input, 
                                          input$tank_cap_max_input, 
                                          by = input$tank_cap_interval_input), 
                             minor_breaks = NULL) + #x axis format
          scale_y_continuous(name = "Efficiency (%)", 
                             breaks = seq(0, 100, by = 20), 
                             limits = c(0, 100),
                             minor_breaks = NULL) + #y axis format
          theme(legend.position = "right",
                #legend.position=c(0.8, 0.5),
                legend.background = element_rect(fill=alpha('white', 0.75))) +
          scale_color_discrete(name = "Efficiency", 
                               labels = c("Storage", "Water-Saving")) +
          theme(text = element_text(color ="grey20", 
                                    size = 12)),
        tooltip = "text")
    
    WS_Sto_chart
    
  }
  
  
  ##Water-saving and storage efficiency for excel
  ws_sto_plot2 <- function(){
    
    #efficiency data frame
    WS_Sto_eff2 <- tangki_nahrim()[['WS_Sto_eff_chart']]
    
    WS_Sto_chart <-
      
      
        ggplot(WS_Sto_eff2, aes(x = Tank_capacity, y = value, group = eff)) +
          geom_line(aes(color = eff), size = 1) +
          geom_point(aes(color = eff,
                         text = hovertext), 
                     size = 2) +
          scale_color_manual(values = c("slategrey", "turquoise3")) +
          theme_bw(base_size = 10) +
          scale_x_continuous(name = expression(paste("Tank Capacity (", m^3, " )")), 
                             breaks = seq(input$tank_cap_min_input, 
                                          input$tank_cap_max_input, 
                                          by = input$tank_cap_interval_input), 
                             minor_breaks = NULL) + #x axis format
          scale_y_continuous(name = "Efficiency (%)", 
                             breaks = seq(0, 100, by = 20), 
                             limits = c(0, 100),
                             minor_breaks = NULL) + #y axis format
          theme(legend.position = "right",
                #legend.position=c(0.8, 0.5),
                legend.background = element_rect(fill=alpha('white', 0.75))) +
          scale_color_discrete(name = "Efficiency", 
                               labels = c("Storage", "Water-Saving")) +
          theme(text = element_text(color ="grey20", 
                                    size = 12),
                plot.caption = element_text(color = "gray40")) +
      labs(caption = gen_info$notes)
        
    
    WS_Sto_chart
    
  }
  
  
  ##Rain and no rain days by year
  raind_yr_plot <- function(){
    
    raind_yr <- tangki_nahrim()[['raind_yr_chart']]
    
    r_d_yr <-
      ggplot(raind_yr, 
             aes(x = year, 
                 y = no_day, 
                 fill=factor(type))) + 
      geom_bar(stat = "identity") +
      scale_fill_manual(values=c("salmon", #color
                                 "steelblue"),
                        name="Legend", 
                        labels = c("No Rain",
                                   "Rain")) +
      theme_bw(base_size = 10) +
      scale_x_continuous(name= "Year",
                         breaks = seq(rain_summary_val$min_year, 
                                      rain_summary_val$max_year, 
                                      by = 1), 
                         minor_breaks = NULL) + #x axis format
      scale_y_continuous(name= "Number of Days",
                         breaks = seq(0, 370, by = 50), 
                         minor_breaks = NULL) + #y axis format
      theme(text=element_text(color="grey20"),
            panel.grid.major.x = element_blank(),
            axis.text.x = element_text(angle = rain_summary_val$angle_precip, 
                                       hjust = 0.5),
            plot.caption = element_text(color = "gray40")) +
      labs(caption = gen_info$notes)
    
    r_d_yr
    
  }
  
  
  
  ##Rain and no rain days by month
  raind_mth_plot <- function(){
    
    raind_mth <- tangki_nahrim()[['raind_mth_chart']]
    
    r_d_mth <-
      ggplot(raind_mth, 
             aes(x = month, y = no_day, 
                 fill=factor(type))) + 
      geom_bar(stat = "identity") +
      scale_fill_manual(values=c("salmon", #color
                                 "steelblue"),
                        name="Legend", 
                        labels = c("No Rain",
                                   "Rain")) +
      theme_bw(base_size = 10) +
      scale_x_continuous(name= "Month",
                         breaks = seq(1, 12, by = 1), 
                         labels=month.abb,
                         minor_breaks = NULL) + #x axis format
      scale_y_continuous(name= "Number of Days",
                         breaks = seq(0, 35, by = 5), 
                         minor_breaks = NULL) + #y axis format
      theme(text=element_text(color="grey20"),
            panel.grid.major.x = element_blank(),
            plot.caption = element_text(color = "gray40")) +
      labs(caption = gen_info$notes)

    r_d_mth
    
  }
  
  
  
  ##Yield-spill by volume
  ys_vol_plot <- function(){
    
    ys_vol <- tangki_nahrim()[['vol_yield_spill_chart']]
    
    ys_v <-
      ggplot(ys_vol,
             aes(x = Tank_capacity, y = volume, 
                 fill=YS)) +
      geom_bar(stat = "identity", 
               position = position_dodge()) +
      #geom_text(aes(label=sprintf("%0.1f", volume)), 
      #          position = position_dodge(width=0.9),vjust=-0.3,size=3,
      #          fontface = 1,color="grey20") + #label data
      scale_fill_manual(values=c("tan1", "mediumturquoise"), 
                        name="Type", 
                        labels = c("Spillage", "Yield")) +
      theme_bw(base_size = 10) +
      scale_x_continuous(name= expression(paste("Tank Capacity ( ", m^3, " )")),
                         breaks = seq(input$tank_cap_min_input, 
                                      input$tank_cap_max_input, 
                                      by = input$tank_cap_interval_input), 
                         minor_breaks = NULL) + #x axis format
      scale_y_continuous(name=expression(paste("Volume per Year ( ", m^3, " )")),
                         #breaks = seq(0, 3000, by = 500), 
                         minor_breaks = NULL) + #y axis format
      theme(text=element_text(color="grey20"),
            panel.grid.major.x = element_blank(),
            plot.caption = element_text(color = "gray40")) +
      labs(caption = gen_info$notes)
    ys_v
    
    
  }
  
  
  ##Yield-spill by day
  ys_day_plot <- function(){
    
    ys_day <- tangki_nahrim()[['day_yield_spill_chart']]
    
    ys_d <-
      ggplot(ys_day,
             aes(x = Tank_capacity, y = No_day, 
                 fill=YS)) +
      geom_bar(stat = "identity", 
               position = position_dodge()) +
      #geom_text(aes(label=sprintf("%0.0f", No_day)), 
      #          position = position_dodge(width=0.9),vjust=-0.3,size=3,
                #family="Roboto Light", fontface = 1, color="grey20") + #label data
      scale_fill_manual(values=c("tan1", "mediumturquoise"), 
                        name="Type", 
                        labels = c("Spillage", "Yield")) +
      theme_bw(base_size = 10) +
      scale_x_continuous(name= expression(paste("Tank Capacity ( ", m^3, " )")),
                         breaks = seq(input$tank_cap_min_input, 
                                      input$tank_cap_max_input, 
                                      by = input$tank_cap_interval_input), 
                         minor_breaks = NULL) + #x axis format
      scale_y_continuous(name=expression(paste("Number of Days per Year")),
                         breaks = seq(0, 380, by = 50), 
                         minor_breaks = NULL) + #y axis format
      theme(text=element_text(color="grey20"),
            panel.grid.major.x = element_blank(),
            plot.caption = element_text(color = "gray40")) +
      labs(caption = gen_info$notes)
    
    ys_d
    
  }
  
  
  ##Tank volume percentage
  tank_vol_plot <- function(){
    
    t_vol <- tangki_nahrim()[['tank_vol_chart']]
    
    legend_color <- c("steelblue3", "steelblue2","lightskyblue","pink1","red")
    
    t_v <- ggplot(t_vol,aes(x = Tank_capacity, y = perc, fill=label_vol,text=hovertext)) +
      scale_fill_manual(values=c("steelblue3", "steelblue2","lightskyblue","pink1","red"),name="Tank Volume") +
      theme_bw(base_size = 10) +
      scale_x_continuous(name= "Tank Capacity (m<sup>3</sup>)",breaks = seq(1,10,by = 1),minor_breaks = NULL) + #x axis format
      scale_y_continuous(name="Percentage Time (%)",breaks = seq(0, 100, by = 10),minor_breaks = NULL) + #y axis format
      theme(text = element_text(color="grey20"),
            panel.grid.major.x = element_blank(),
            legend.position = "right")
    
    t_v <- t_v + 
      geom_bar(position = "stack", stat = "identity")
    
    
    t_v_fig <- ggplotly(t_v, tooltip = "text")
    
    t_v_fig
    
  }
  
  
  ##Tank volume percentage for excel
  tank_vol_plot2 <- function(){
    
    t_vol <- tangki_nahrim()[['tank_vol_chart']]
    
    t_v2 <-
      
      ggplot(t_vol,
             aes(x = Tank_capacity, y = perc, 
                 fill=factor(vol, 
                             levels=c("tankvol_100", #reorder 
                                      "tankvol_75",
                                      "tankvol_50",
                                      "tankvol_25",
                                      "tankvol_0")))) +
      geom_bar(stat = "identity") +
      scale_fill_manual(values=c("steelblue3", #color
                                 "steelblue2",
                                 "lightskyblue",
                                 "pink1",
                                 "red"), 
                        name="Percentage Tank Volume", 
                        labels = c("Tank Volume 75 - 100%",
                                   "Tank Volume 50 - 75%", 
                                   "Tank Volume 25 - 50%",
                                   "Tank Volume less than 25%",
                                   "Tank Volume Empty")) +
      theme_bw(base_size = 10) +
      scale_x_continuous(name= expression(paste("Tank Capacity ( ", m^3, " )")),
                         breaks = seq(input$tank_cap_min_input, 
                                      input$tank_cap_max_input, 
                                      by = input$tank_cap_interval_input), 
                         minor_breaks = NULL) + #x axis format
      scale_y_continuous(name=expression(paste("Percentage Time (%)")),
                         breaks = seq(0, 100, by = 10), 
                         minor_breaks = NULL) + #y axis format
      theme(text = element_text(color="grey20"),
            panel.grid.major.x = element_blank(),
            #legend.position=c(0.77, 0.5),
            legend.position = "right",
            legend.background = element_rect(fill=alpha('white', 0.75)),
            plot.caption = element_text(color = "gray40")) +
      labs(caption = gen_info$notes)
    
    t_v2
    
  }
  
  
  
  
  
  #CALCULATE BUTTON
  
  output$Calculate_action <- renderPrint({ input$Calculate_action })
  
  observeEvent(input$Calculate_action,{
    
    
    #Plot efficiency charts
    output$WS_Sto_eff_Plot <- renderPlotly({
      
      ws_sto_plot()

    })
    
    #show efficiency results table
    #output$WS_Sto_eff_Table <- renderTable({
      
      #table for presentation
      #tangki_nahrim()[['WS_Sto_eff_table']]
      
    #}, 
    
    #table properties
    #rownames = FALSE,
    #hover = TRUE,
    #align = '?', 
    #digits = 1
    
    #)
    
    
    #Plot tank volume percentage
    output$tank_volume_plot <- renderPlotly({
      
      tank_vol_plot()
      
    })
    
    
    #switch to tab
    updateTabsetPanel(session, "inTabset",selected = "Results")
    #js$toTop();
    
    

  }
  
  )
  
  #TEXT RESULTS
  #output$results_text <- renderText()
  
  
  #DOWNLOAD XLSX OF RESULTS
  output$downloadResults <- downloadHandler(
    
    # filename save as
    filename = function() {
      paste0("TangkiNAHRIM_RF",
             input$stn_Output, 
             "_A", (input$roof_length_input*input$roof_width_input), 
             "_D", input$water_demand_input, ".xlsx")
    },
    
    # content of xlsx
    content = function(file) {
      
      
      
      # input data
      ## tank capacity iteration
      Tank_capacity_1 <- as.numeric(input$tank_cap_min_input) #in m3, first value of tank capacity
      Tank_capacity_10 <- as.numeric(input$tank_cap_max_input) #in m3, last value of tank capacity
      Tank_capacity_interval <- as.numeric(input$tank_cap_interval_input) #in m3, tank capacity interval
      tank_mul <- (Tank_capacity_10 - Tank_capacity_1)/Tank_capacity_interval + 1
      ## year with data
      year_w_data <- rain_summary_val$yr_w_data
      
      
      
      # summary page
      input_param <- c('Station Number', 
                       'Length of record (year)', 
                       'Length of record (day)',
                       #'Days with no rain (%)',
                       'Roof area (sqm)',
                       'Roof coefficient',
                       'First flush (mm)',
                       'Water demand (litres per day)',
                       'Tank size range (cubic meter)')
      input_value <- c(input$stn_Output,
                       rain_summary_val$yr_w_data,
                       rain_summary_val$day_w_data,
                       #lapply(day_norain/day_w_data*100, round, digits = 1),
                       (input$roof_length_input*input$roof_width_input),
                       input$runoff_coef_input,
                       input$first_flush_input,
                       input$water_demand_input,
                       paste0(input$tank_cap_min_input, " to ", input$tank_cap_max_input))
      input_value2 <- c(input$stn_Output,
                       rain_summary_val$yr_w_data,
                       rain_summary_val$day_w_data,
                       #lapply(day_norain/day_w_data*100, round, digits = 1),
                       (input$roof_length_input*input$roof_width_input),
                       input$runoff_coef_input,
                       input$first_flush_input,
                       input$water_demand_input,
                       paste0(input$tank_cap_min_input, " to ", input$tank_cap_max_input))
      input_value3 <- c(filtered()$STATION_NA,
                        "",
                        "",
                        #"",
                        "",
                        "",
                        "",
                        "",
                        "")
      input_page <- data.frame(cbind(input_param, input_value, input_value2, input_value3))
      
      

      # Create a blank workbook
      wb <- createWorkbook(creator = "Water Resources Unit, NAHRIM", title = gen_info$version_no)
      addWorksheet(wb, "Report")
      
      
      
      # write all in first sheet
      ## start row for every table
      startRow_precip_yr <- ifelse(15 + tank_mul < 38, 
                                   38,
                                   15 + tank_mul + 3)
      startRow_precip_mth <- ifelse(startRow_precip_yr + 2 + year_w_data < startRow_precip_yr + 26, 
                                    startRow_precip_yr + 26,
                                    startRow_precip_yr + 2 + year_w_data + 3)
      startRow_raind_yr <- startRow_precip_mth + 25
      startRow_raind_mth <- ifelse(startRow_raind_yr + 2 + year_w_data < startRow_raind_yr + 26, 
                                   startRow_raind_yr + 26,
                                   startRow_raind_yr + 2 + year_w_data + 3)
      startRow_ys_vol <- startRow_raind_mth + 25
      startRow_ys_day <- ifelse(startRow_ys_vol + 1 + tank_mul < startRow_ys_vol + 25, 
                                startRow_ys_vol + 25,
                                startRow_ys_vol + 1 + tank_mul + 3)
      startRow_tankvol <- ifelse(startRow_ys_day + 1 + tank_mul < startRow_ys_day + 25, 
                                 startRow_ys_day + 25,
                                 startRow_ys_day + 1 + tank_mul + 3)
      ### all rows for heading 1
      df_row_head1 <- c(14, 
                        startRow_precip_yr, 
                        startRow_raind_yr, 
                        startRow_ys_vol, 
                        startRow_ys_day,
                        startRow_tankvol)
      ### all rows for heading 2
      df_row_head2 <- c(startRow_precip_yr + 1,
                        startRow_precip_mth,
                        startRow_raind_yr + 1,
                        startRow_raind_mth)
      ### all rows notes
      df_row_notes <- c(1, 3, 7)
      ### all rows table column names
      df_row_table_name <- c(15,
                             startRow_precip_yr + 2,
                             startRow_precip_mth + 1,
                             startRow_raind_yr + 2,
                             startRow_raind_mth + 1,
                             startRow_ys_vol + 1, 
                             startRow_ys_day + 1,
                             startRow_tankvol + 1)
      ### all rows table bottom
      df_row_table_bottom <- c(15 + tank_mul,
                               startRow_precip_yr + 2 + year_w_data,
                               startRow_precip_mth + 1 + 12,
                               startRow_raind_yr + 2 + year_w_data,
                               startRow_raind_mth + 1 + 12,
                               startRow_ys_vol + 1 + tank_mul, 
                               startRow_ys_day + 1 + tank_mul,
                               startRow_tankvol + 1 + tank_mul)
      ### each table column number
      df_col_table <- c(3,
                        2, 2,
                        3, 3,
                        3, 3,
                        6)
      ### page break rows
      df_row_page <- c(startRow_precip_yr - 1,
                       startRow_raind_yr - 1,
                       startRow_ys_vol - 1, 
                       startRow_tankvol- 1)
      
      
      
      # write data
      writeData(wb, "Report", "Tangki NAHRIM Calculation",
                startCol = 1, startRow = 1)
      writeData(wb, "Report", paste0("Generated by ", gen_info$version_no, " on ", Sys.Date()),
                startCol = 8, startRow = 1)
      #writeData(wb, "Report", "Reference:",
      #          startCol = 8, startRow = 3)
      writeData(wb, "Report", "Website:",
                startCol = 8, startRow = 7)
      writeData(wb, "Report", "https://waterresources-nahrim.shinyapps.io/Tangki_NAHRIM/",
                startCol = 8, startRow = 8)
      writeData(wb, "Report", input_page,
                startCol = 1, startRow = 2)
      writeData(wb, "Report", "Water-Saving and Storage Efficiency",
                startCol = 1, startRow = 14)
      writeData(wb, "Report", tangki_nahrim()[['WS_Sto_eff_table']],
                startCol = 1, startRow = 15)
      writeData(wb, "Report", "Rainfall Data",
                startCol = 1, startRow = startRow_precip_yr)
      writeData(wb, "Report", "Annual Rainfall",
                startCol = 1, startRow = startRow_precip_yr + 1)
      writeData(wb, "Report", rain_summary()[['rain_yr_table']],
                startCol = 1, startRow = startRow_precip_yr + 2)
      writeData(wb, "Report", "Monthly Rainfall",
                startCol = 1, startRow = startRow_precip_mth)
      writeData(wb, "Report", rain_summary()[['rain_mth_table']],
                startCol = 1, startRow = startRow_precip_mth + 1)
      writeData(wb, "Report", "Rain days vs No Rain days",
                startCol = 1, startRow = startRow_raind_yr)
      writeData(wb, "Report", "Annual Rain days vs No Rain days",
                startCol = 1, startRow = startRow_raind_yr + 1)
      writeData(wb, "Report", tangki_nahrim()[['raind_yr_table']],
                startCol = 1, startRow = startRow_raind_yr + 2)
      writeData(wb, "Report", "Monthly Rain days vs No Rain days",
                startCol = 1, startRow = startRow_raind_mth)
      writeData(wb, "Report", tangki_nahrim()[['raind_mth_table']],
                startCol = 1, startRow = startRow_raind_mth + 1)
      writeData(wb, "Report", "Yield vs Spillage by Volume",
                startCol = 1, startRow = startRow_ys_vol)
      writeData(wb, "Report", tangki_nahrim()[['vol_yield_spill_table']],
                startCol = 1, startRow = startRow_ys_vol + 1)
      writeData(wb, "Report", "Yield vs Spillage by Day",
                startCol = 1, startRow = startRow_ys_day)
      writeData(wb, "Report", tangki_nahrim()[['day_yield_spill_table']],
                startCol = 1, startRow = startRow_ys_day + 1)
      writeData(wb, "Report", "Percentage Tank Volume",
                startCol = 1, startRow = startRow_tankvol)
      writeData(wb, "Report", tangki_nahrim()[['tank_vol_table']],
                startCol = 1, startRow = startRow_tankvol + 1)
      
      
      
      ## set col widths
      setColWidths(wb, "Report", cols = 1, widths = 15.33)
      setColWidths(wb, "Report", cols = 2:6, widths = 14)
      
      
      ## print and save charts
      chart_res <- 300
      
      print(rain_yr_plot())
      ggsave(paste0(tempdir(), "/", "rain_yr_plot.png"), dpi = chart_res,
             width = 18.9, height = 11.58, units = "cm")
      
      print(rain_mth_plot())
      ggsave(paste0(tempdir(), "/", "rain_mth_plot.png"), dpi = chart_res,
             width = 18.9, height = 11.58, units = "cm")
      
      #print(ws_sto_plot())
      #ggsave(paste0(tempdir(), "/", "ws_sto_plot.png"), dpi = chart_res)
      
      print(ws_sto_plot2())
      ggsave(paste0(tempdir(), "/", "ws_sto_plot2.png"), dpi = chart_res,
             width = 18.9, height = 11.58, units = "cm")
      
      print(raind_yr_plot())
      ggsave(paste0(tempdir(), "/", "raind_yr_plot.png"), dpi = chart_res,
             width = 18.9, height = 11.58, units = "cm")
      
      print(raind_mth_plot())
      ggsave(paste0(tempdir(), "/", "raind_mth_plot.png"), dpi = chart_res,
             width = 18.9, height = 11.58, units = "cm")
      
      print(ys_vol_plot())
      ggsave(paste0(tempdir(), "/", "ys_vol_plot.png"), dpi = chart_res,
             width = 18.9, height = 11.58, units = "cm")
      
      print(ys_day_plot())
      ggsave(paste0(tempdir(), "/", "ys_day_plot.png"), dpi = chart_res,
             width = 18.9, height = 11.58, units = "cm")
      
      print(tank_vol_plot2())
      ggsave(paste0(tempdir(), "/", "tank_vol_plot.png"), dpi = chart_res,
             width = 18.9, height = 11.58, units = "cm")

      
      ## insert charts into worksheets
      insertImage(wb, "Report", paste0(tempdir(), "/", "ws_sto_plot2.png"), 
                  width = 18.9, height = 11.58, 
                  startRow = 14, startCol = 5, units = "cm")
      insertImage(wb, "Report", paste0(tempdir(), "/", "rain_yr_plot.png"), 
                  width = 18.9, height = 11.58, 
                  startRow = startRow_precip_yr + 2, startCol = 4, units = "cm")
      insertImage(wb, "Report", paste0(tempdir(), "/", "rain_mth_plot.png"), 
                  width = 18.9, height = 11.58, 
                  startRow = startRow_precip_mth + 1, startCol = 4, units = "cm")
      insertImage(wb, "Report", paste0(tempdir(), "/", "raind_yr_plot.png"), 
                  width = 18.9, height = 11.58, 
                  startRow = startRow_raind_yr + 2, startCol = 5, units = "cm")
      insertImage(wb, "Report", paste0(tempdir(), "/", "raind_mth_plot.png"), 
                  width = 18.9, height = 11.58, 
                  startRow = startRow_raind_mth + 1, startCol = 5, units = "cm")
      insertImage(wb, "Report", paste0(tempdir(), "/", "ys_vol_plot.png"), 
                  width = 18.9, height = 11.58, 
                  startRow = startRow_ys_vol + 1, startCol = 5, units = "cm")
      insertImage(wb, "Report", paste0(tempdir(), "/", "ys_day_plot.png"), 
                  width = 18.9, height = 11.58, 
                  startRow = startRow_ys_day + 1, startCol = 5, units = "cm")
      insertImage(wb, "Report", paste0(tempdir(), "/", "tank_vol_plot.png"), 
                  width = 18.9, height = 11.58, 
                  startRow = startRow_tankvol + 1 + tank_mul + 3, startCol = 1, units = "cm")
      
      
      
      # formatting
      ##  create styles
      ###  title
      title_style <- createStyle(fontName = "Tahoma", fontSize = 20, fontColour = "navyblue",
                                 textDecoration = "bold")
      ### heading 1
      head1_style <- createStyle(fontSize = 14, fontColour = "navyblue",
                                 textDecoration = "bold")
      ### heading 2
      head2_style <- createStyle(fontSize = 12, fontColour = "navyblue")
      ### input parameters
      input_style <- createStyle(textDecoration = "bold")
      ### notes
      note_style <- createStyle(textDecoration = "italic")
      ### title top line
      line_style <- createStyle(border = "bottom", borderColour = "black", borderStyle = "thin")
      ### table - column names
      table_name_style <- createStyle(border = "TopBottom", borderColour = "black", 
                                      borderStyle = "thin", fgFill = "gray85",
                                      textDecoration = "bold", wrapText = T,
                                      halign = "center", valign = "center")
      ### table - bottom
      table_bottom_style <- createStyle(border = "Bottom", borderColour = "black", 
                                        borderStyle = "thin")
      
      
      ## apply styles
      ### title
      addStyle(wb, "Report", style = title_style, stack = T,
               rows = 1, cols = 1)
      ### all heading 1
      for (i in df_row_head1){
        addStyle(wb, "Report", style = head1_style, stack = T,
                 rows = i, cols = 1)
      }
      ### all heading 2
      for (j in df_row_head2){
        addStyle(wb, "Report", style = head2_style, stack = T,
                 rows = j, cols = 1)
      }
      ### input parameters
      addStyle(wb, "Report", style = input_style, stack = T,
               rows = 3:11, cols = 1, gridExpand = T)
      ### all notes
      for (k in df_row_notes){
        addStyle(wb, "Report", style = note_style, stack = T,
                 rows = k, cols = 8)
      }
      ### title top line
      addStyle(wb, "Report", style = line_style, stack = T,
               rows = 1, cols = 1:14, gridExpand = T)
      ### table - column names
      mapply(function(a, b) addStyle(wb, "Report", style = table_name_style, stack = T,
                                     rows = a, cols = 1:b, gridExpand = T), 
             df_row_table_name,
             df_col_table)
      ### table - bottom
      mapply(function(c, d) addStyle(wb, "Report", style = table_bottom_style, stack = T,
                                     rows = c, cols = 1:d, gridExpand = T), 
             df_row_table_bottom,
             df_col_table)
      
      
      # cleaning up
      deleteData(wb, "Report", rows = 3:11, cols = 2, gridExpand = T)
      deleteData(wb, "Report", rows = 2, cols = 1:4, gridExpand = T)
      
      
      # page break
      for (x in df_row_page){
        pageBreak(wb, "Report", i = x, type = "row")
      }
      pageBreak(wb, "Report", i = 14, type = "column")
      
      
      # page setup
      #pageSetup(wb, "Report", orientation = "landscape", fitToWidth = T)
      
      
      saveWorkbook(wb, 
                   file, 
                   overwrite = TRUE)
      
    }
  )

  
}


shinyApp(ui = ui, server = server)