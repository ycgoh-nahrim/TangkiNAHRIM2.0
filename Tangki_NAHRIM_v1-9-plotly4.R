# load packages
library(tidyverse)
library(lubridate)
library(scales)
library(extrafont)
library(openxlsx)
library(plotly)


# set strings as factors to false
options(stringsAsFactors = FALSE)



#FOR Tangki NAHRIM CALCULATION FROM RAINFALL DATA##############################################

#INPUT

#parameters
station_no <- "3014091" #station number
roof_length <- 10 #in meters
roof_width <- 10 #in meters
Runoff_coef <- 0.8 #usually 0.8 for concrete tiles, 0.9 for zink/metals
First_flush <- 1 #in mm
Water_demand_l <- 200 #in litres
Tank_capacity_1 <- 1 #in m3, first value of tank capacity
Tank_capacity_10 <- 10 #in m3, last value of tank capacity
Tank_capacity_interval <- 1 #in m3, tank capacity interval

#############

#set working directory
#set filename
filename2 <- paste0(station_no, "_A", roof_length*roof_width, "_D", Water_demand_l)
# get current working directory
working_dir <- getwd()
dir.create(filename2)
setwd(filename2)



##OPTIONS (rainfall station data or from database) #############
#input format in c(Date, Depth) in csv (case sensitive), station_no is station number, Depth in mm
station_no <- strsplit("3014091.csv", "\\.")[[1]][1]
raindata <- read.csv(file=paste0(working_dir,"/", station_no, ".csv"),header = TRUE,sep=",") #depends on working dir
raindata$Date <- as.Date(raindata$Date, format = "%d/%m/%Y")

#from rainfall database
rain_db <- read.csv(file="F:/Documents/2019/20160506 Tangki NAHRIM/Data/SQL/rainfall_pm_clean6d.csv",
                    header = TRUE, sep=",")
raindata <- rain_db %>%
  filter(stn_no == station_no) %>% #check field name
  select(date, depth) %>%
  arrange(date)
raindata <- raindata %>% 
  rename(Date = date, Depth = depth)
#raindata$Date <- as.Date(raindata$Date, format = "%d/%m/%Y")
raindata$Date <- as.Date(raindata$Date, format = "%Y-%m-%d")


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

#find max and min year
max_year <- max(tankdata$year)
min_year <- min(tankdata$year)
total_year <- max_year - min_year + 1
year_w_data <- length(unique(tankdata$year))
day_w_data <- length(unique(tankdata$Date))
angle_precip <- ifelse(total_year > 20, 90, 0)


#CALCULATION FOR ANALYSIS##############################################

# calculate the sum precipitation for each year
precip_sum_yr <- tankdata %>%
  group_by(year) %>%
  summarise(sum_precip = sum(Depth)/tank_mul) #tank sizes iteration

# calculate the sum precipitation for each month, each year
precip_sum_mth_yr <- tankdata %>%
  group_by(month, year) %>%
  summarise(sum_precip = sum(Depth)/tank_mul) #tank sizes iteration
precip_sum_mth <- precip_sum_mth_yr %>%
  group_by(month) %>%
  summarise(mth_precip = mean(sum_precip))


# calculate the sum rain days and no rain days for each year
raind_yr <- tankdata %>%
  group_by(year) %>%
  summarise(sum_rain = sum(Runoff > 0)/tank_mul, 
            sum_norain = sum(Runoff == 0)/tank_mul) #tank sizes iteration
#pivot for stacked bar plot
raind_yr2 <- raind_yr %>%
  gather(type, no_day, sum_rain:sum_norain)

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

#format number
WS_Sto_eff4 <- WS_Sto_eff
WS_Sto_eff4$WS_eff <- lapply(WS_Sto_eff4$WS_eff, round, digits = 1)
WS_Sto_eff4$Sto_eff <- lapply(WS_Sto_eff4$Sto_eff, round, digits = 1)
#table for output
WS_Sto_eff4 <- WS_Sto_eff4 %>%
  rename("Water-Saving Efficiency" = WS_eff) %>%
  rename("Storage Efficiency" = Sto_eff)
#add columns for hover text
WS_Sto_eff2 <- WS_Sto_eff2 %>%
  mutate(hovertext = ifelse(WS_Sto_eff2$eff == 'WS_eff',
                            paste0('<b>',
                                   sprintf("%0.1f", value),
                                   '%</b> of water demand can be met for tank size <b>',
                                   Tank_capacity,
                                   ' m<sup>3</sup></b><br>'),
                            paste0('<b>',
                                   sprintf("%0.1f", value),
                                   '%</b> of rainwater from roof can be utilized for tank size <b>',
                                   Tank_capacity,
                                   ' m<sup>3</sup></b><br>')))



#calculate percentage tank volume
tank_vol <- tankdata %>%
  group_by(Tank_capacity) %>%
  summarise(tankvol_0 = sum(Tank_0)/n()*100, 
            tankvol_25 = sum(Tank_25)/n()*100,
            tankvol_50 = sum(Tank_50)/n()*100,
            tankvol_75 = sum(Tank_75)/n()*100,
            tankvol_100 = sum(Tank_100)/n()*100)
#pivot for stacked bar plot
tank_vol2 <- tank_vol %>%
  gather(vol, perc, tankvol_0:tankvol_100)


#CHARTS##############################################

#bar chart, annual rainfall
#Total Annual Precipitation
precip_sum_yr_chart <- precip_sum_yr %>%
  ggplot(aes(x = year, y = sum_precip)) +
  geom_bar(stat = "identity", fill = "steelblue") +
  #geom_text(aes(label=sum_precip), position=position_dodge(width=0.9), vjust=-0.3, size=3.5) + #label data
  theme_bw(base_size = 10) +
  scale_x_continuous(name= "Year", 
                     breaks = seq(min_year, max_year, by = 1), 
                     minor_breaks = NULL) + #x axis format
  scale_y_continuous(name= "Annual precipitation (mm)",
                     #breaks = seq(0, 3500, by = 500), 
                     minor_breaks = NULL) + #y axis format
  geom_hline(aes(yintercept = mean(sum_precip)), 
             color="black", 
             alpha=0.3, 
             size=1) + #avg line
  geom_text(aes(min_year+1,
                mean(sum_precip),
                label = paste("Average = ", sprintf("%0.0f", mean(sum_precip)), "mm"), 
                vjust = -0.5, hjust = 0), 
            size=3.5, 
            family="Roboto Light", 
            fontface = 1, 
            color="grey20") + #avg line label
  theme(text=element_text(family="Roboto", color="grey20"),
        panel.grid.major.x = element_blank(),
        axis.text.x = element_text(angle = angle_precip, hjust = 0.5))
ggplotly(precip_sum_yr_chart)
#plot_ly(precip_sum_yr_chart)
#print last plot to file
ggsave(paste0(filename2, "_precip_annual.jpg"), dpi = 300)


#bar chart, monthly rainfall
#Long-term Average Monthly Precipitation
precip_sum_mth %>%
  ggplot(aes(x = month, y = mth_precip)) +
  geom_bar(stat = "identity", fill = "steelblue") +
  theme_bw(base_size = 10) +
  scale_x_continuous(name = "Month",
                     breaks = seq(1, 12, by = 1), 
                     minor_breaks = NULL, 
                     labels=month.abb) + #x axis format
  scale_y_continuous(name = "Monthly precipitation (mm)",
                     #breaks = seq(0, 350, by = 50), 
                     minor_breaks = NULL) + #y axis format
  theme(text=element_text(family="Roboto", 
                          color="grey20"),
        panel.grid.major.x = element_blank())
#print last plot to file
ggsave(paste0(filename2, "_precip_mth.jpg"), dpi = 300)


#stacked bar chart, yearly rain/no rain days
#Rain days versus No Rain days, by Year
raind_yr2 %>%
  ggplot(aes(x = year, y = no_day, 
             fill=factor(type))) + 
  geom_bar(stat = "identity") +
  scale_fill_manual(values=c("salmon", #color
                             "steelblue"),
                    name="Legend", 
                    labels = c("No Rain",
                               "Rain")) +
  theme_bw(base_size = 10) +
  scale_x_continuous(name= "Year",
                     breaks = seq(min_year, max_year, by = 1), 
                     minor_breaks = NULL) + #x axis format
  scale_y_continuous(name= "Number of Days",
                     breaks = seq(0, 370, by = 50), 
                     minor_breaks = NULL) + #y axis format
  theme(text=element_text(family="Roboto", 
                          color="grey20", 
                          size = 12),
        panel.grid.major.x = element_blank(),
        axis.text.x = element_text(angle = angle_precip, hjust = 0.5))
#print last plot to file
ggsave(paste0(filename2, "_rain_yr.jpg"), dpi = 300)


#stacked bar chart, monthly rain/no rain days
#Rain days versus No Rain days, by Month
raind_mth2 %>%
  ggplot(aes(x = month, y = no_day, 
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
  theme(text=element_text(family="Roboto", 
                          color="grey20", 
                          size = 12),
        panel.grid.major.x = element_blank())
#print last plot to file
ggsave(paste0(filename2, "_rain_mth.jpg"), dpi = 300)



#line chart, WS & Storage efficiency
#Performance Measures
WS_Sto_eff2 %>%
  ggplot(aes(x = Tank_capacity, y = value, group = eff)) +
  geom_line(aes(color = eff), size = 1) +
  geom_point(aes(color = eff), size = 2) +
  scale_color_manual(values=c("slategrey", "turquoise3")) +
  geom_text(aes(label=sprintf("%0.1f", value)), 
            position = position_dodge(width=0.9), 
            vjust = -0.3, 
            size = 3.5,
            family ="Roboto Light", 
            fontface = 1, 
            color = "grey20") + #label data
  theme_bw(base_size = 10) +
  scale_x_continuous(name = expression(paste("Tank Capacity (", m^3, " )")), 
                     breaks = seq(Tank_capacity_1, Tank_capacity_10, by = Tank_capacity_interval), 
                     minor_breaks = NULL) + #x axis format
  scale_y_continuous(name = "Efficiency (%)", 
                     breaks = seq(0, 100, by = 10), 
                     minor_breaks = NULL) + #y axis format
  theme(legend.position=c(0.8, 0.5),
        legend.background = element_rect(fill=alpha('white', 0.75))) +
  scale_color_discrete(name="Efficiency", 
                       labels = c("Storage Efficiency", "Water-Saving Efficiency")) +
  theme(text=element_text(family="Roboto", 
                          color="grey20", 
                          size = 12))
#print last plot to file
ggsave(paste0(filename2, "_eff.jpg"), dpi = 300)


#bar chart, volume - yield and spillage
#Rainwater Yield and Spillage, by Volume
vol_yield_spill2 %>%
  ggplot(aes(x = Tank_capacity, y = volume, fill=YS)) +
  geom_bar(stat = "identity", 
           position = position_dodge()) +
  geom_text(aes(label=sprintf("%0.1f", volume)), 
            position = position_dodge(width=0.9), 
            vjust=-0.3, 
            size=3,
            family="Roboto Light", 
            fontface = 1, 
            color="grey20") + #label data
  scale_fill_manual(values=c("tan1", "mediumturquoise"), 
                    name="Legend", 
                    labels = c("Spillage", "Yield")) +
  theme_bw(base_size = 10) +
  scale_x_continuous(name= expression(paste("Tank Capacity (", m^3, " )")),
                     breaks = seq(Tank_capacity_1, Tank_capacity_10, by = Tank_capacity_interval), 
                     minor_breaks = NULL) + #x axis format
  scale_y_continuous(name=expression(paste("Volume per Year (", m^3, " )")),
                     #breaks = seq(0, 3000, by = 500), 
                     minor_breaks = NULL) + #y axis format
  theme(text=element_text(family="Roboto", 
                          color="grey20", 
                          size = 12),
        panel.grid.major.x = element_blank())
#print last plot to file
ggsave(paste0(filename2, "_YS-vol.jpg"), dpi = 300)


#bar chart, day - yield and spillage
#Rainwater Yield and Spillage, by Days
day_yield_spill2 %>%
  ggplot(aes(x = Tank_capacity, y = No_day, fill=YS)) +
  geom_bar(stat = "identity", 
           position = position_dodge()) +
  geom_text(aes(label=sprintf("%0.0f", No_day)), 
            position = position_dodge(width=0.9), 
            vjust=-0.3, 
            size=3,
            family="Roboto Light", 
            fontface = 1, 
            color="grey20") + #label data
  scale_fill_manual(values=c("tan1", "mediumturquoise"), 
                    name="Legend", 
                    labels = c("Spillage", "Yield")) +
  theme_bw(base_size = 10) +
  scale_x_continuous(name= expression(paste("Tank Capacity (", m^3, " )")),
                     breaks = seq(Tank_capacity_1, Tank_capacity_10, by = Tank_capacity_interval), 
                     minor_breaks = NULL) + #x axis format
  scale_y_continuous(name=expression(paste("Number of Days per Year")),
                     breaks = seq(0, 380, by = 50), 
                     minor_breaks = NULL) + #y axis format
  theme(text=element_text(family="Roboto", 
                          color="grey20", 
                          size = 12),
        panel.grid.major.x = element_blank())
#print last plot to file
ggsave(paste0(filename2, "_YS-day.jpg"), dpi = 300)


#bar chart, percentage tank volume
#Percentage Tank Volume
tank_vol2 %>%
  ggplot(aes(x = Tank_capacity, y = perc, 
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
  scale_x_continuous(name= expression(paste("Tank Capacity (", m^3, " )")),
                     breaks = seq(Tank_capacity_1, Tank_capacity_10, by = Tank_capacity_interval), 
                     minor_breaks = NULL) + #x axis format
  scale_y_continuous(name = expression(paste("Percentage Time (%)")),
                     breaks = seq(0, 100, by = 10), 
                     minor_breaks = NULL) + #y axis format
  theme(text=element_text(family="Roboto", 
                          color="grey20", 
                          size = 12),
        panel.grid.major.x = element_blank(),
        legend.position = "right",
        legend.background = element_rect(fill=alpha('white', 0.75)))
#print last plot to file
ggsave(paste0(filename2, "_tankvolper.jpg"), dpi = 300)


## OUTPUT####################
#raw output in csv, selected columns
write.csv(tankdata[c(1:8,10,14:19)], file = paste0("TangkiNAHRIM_", filename2, "_raw.csv"), row.names = FALSE)


#summary xlsx
#replace column names in dataframes
precip_sum_yr2 <- precip_sum_yr %>% 
  rename(Year = year,
         Precipitation_mm = sum_precip)
precip_sum_mth2 <- precip_sum_mth %>% 
  rename(Month = month,
         Precipitation_mm = mth_precip)
raind_yr3 <- raind_yr2 %>% 
  rename(Year = year,
         Type = type,
         Days = no_day)
raind_mth3 <- raind_mth2 %>% 
  rename(Month = month,
         Type = type,
         Days = no_day)
WS_Sto_eff3 <- WS_Sto_eff2 %>% 
  rename(Efficiency = eff,
         Percentage = value)
vol_yield_spill3 <- 
  vol_yield_spill2 %>% 
  rename(Type = YS,
         Volume_m3 = volume)
day_yield_spill3 <- day_yield_spill2 %>% 
  rename(Type = YS,
         Days = No_day)
tank_vol3 <- tank_vol2 %>% 
  rename(Percentage_volume = vol,
         Percentage = perc)

#replace string
raind_yr3$Type <- raind_yr3$Type %>% 
  str_replace("sum_norain", "No Rain") %>%
  str_replace("sum_rain", "Rain")
raind_mth3$Type <- raind_mth3$Type %>% 
  str_replace("mth_norain", "No Rain") %>%
  str_replace("mth_rain", "Rain")
WS_Sto_eff3$Efficiency <- WS_Sto_eff3$Efficiency %>% 
  str_replace("WS_eff", "Water-saving") %>%
  str_replace("Sto_eff", "Storage")
vol_yield_spill3$Type <- vol_yield_spill3$Type %>% 
  str_replace("yield_yr", "Yield") %>%
  str_replace("spill_yr", "Spill")
day_yield_spill3$Type <- day_yield_spill3$Type %>% 
  str_replace("yield_yr", "Yield") %>%
  str_replace("spill_yr", "Spill")
tank_vol3$Percentage_volume <- tank_vol3$Percentage_volume %>% 
  str_replace("tankvol_0", "Empty") %>%
  str_replace("tankvol_25", "< 25%") %>%
  str_replace("tankvol_50", "25% - 50%") %>%
  str_replace("tankvol_75", "50% - 75%") %>%
  str_replace("tankvol_100", "75% - 100%")

#rounding, numbers
precip_sum_yr2$Precipitation_mm <- lapply(precip_sum_yr2$Precipitation_mm, 
                                          round, digits = 1)
precip_sum_yr2$Precipitation_mm <- as.numeric(precip_sum_yr2$Precipitation_mm)
precip_sum_mth2$Precipitation_mm <- lapply(precip_sum_mth2$Precipitation_mm, 
                                           round, digits = 1)
precip_sum_mth2$Precipitation_mm <- as.numeric(precip_sum_mth2$Precipitation_mm)
raind_yr3$Days <- lapply(raind_yr3$Days, round, digits = 0)
raind_yr3$Days <- as.numeric(raind_yr3$Days)
raind_mth3$Days <- lapply(raind_mth3$Days, round, digits = 1)
raind_mth3$Days <- as.numeric(raind_mth3$Days)
WS_Sto_eff4$"Water-Saving Efficiency" <- as.numeric(WS_Sto_eff4$"Water-Saving Efficiency")
WS_Sto_eff4$"Storage Efficiency" <- as.numeric(WS_Sto_eff4$"Storage Efficiency")
vol_yield_spill3$Volume_m3 <- lapply(vol_yield_spill3$Volume_m3, round, digits = 1)
vol_yield_spill3$Volume_m3 <- as.numeric(vol_yield_spill3$Volume_m3)
day_yield_spill3$Days <- lapply(day_yield_spill3$Days, round, digits = 0)
day_yield_spill3$Days <- as.numeric(day_yield_spill3$Days)
tank_vol3$Percentage <- lapply(tank_vol3$Percentage, round, digits = 1)
tank_vol3$Percentage <- as.numeric(tank_vol3$Percentage)

#summary page
input_param <- c('Station Number', 
                 'Length of record (year)', 
                 'Length of record (day)',
                 'Days with no rain (%)',
                 'Roof area (sqm)',
                 'Roof coefficient',
                 'First flush (mm)',
                 'Water demand (litres per day)',
                 'Tank size range (cubic meter)')
input_value <- c(station_no,
                 year_w_data,
                 day_w_data,
                 lapply(day_norain/day_w_data*100, round, digits = 1),
                 (roof_length*roof_width),
                 Runoff_coef,
                 First_flush,
                 Water_demand_l,
                 paste0(Tank_capacity_1, " to ", Tank_capacity_10))
input_page <- data.frame(cbind(input_param, input_value))
input_page <- input_page %>%
        rename("Tangki NAHRIM Calculation" = input_param,
               "Input" = input_value)


#list all dataframe
list_worksheet <- list("Input" = input_page,
                       "Precip_yr" = precip_sum_yr2,
                       "Precip_mth" = precip_sum_mth2,
                       "WS_Storage Efficiency" = WS_Sto_eff4,
                       "Rain_norain_yr" = raind_yr3,
                       "Rain_norain_mth" = raind_mth3,                        
                       "Yield_Spill (Volume)" = vol_yield_spill3, 
                       "Yield_Spill (Day)" = day_yield_spill3,
                       "Tank volume" = tank_vol3)

# Create a blank workbook
wb <- createWorkbook()

# Loop through the list of split tables as well as their names
#   and add each one as a sheet to the workbook
Map(function(data, name){
  addWorksheet(wb, name)
  writeData(wb, name, data)
  setColWidths(wb, name, cols = 1:3, widths = "auto")
}, list_worksheet, names(list_worksheet))

## set col widths
setColWidths(wb, "Precip_yr", cols = 1, widths = 5)
setColWidths(wb, "Rain_norain_yr", cols = 1, widths = 5)

#insert image
insertImage(wb, "Precip_yr", paste0(filename2, "_precip_annual.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 2, startCol = 5, units = "cm")
insertImage(wb, "Precip_mth", paste0(filename2, "_precip_mth.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 2, startCol = 5, units = "cm")
insertImage(wb, "WS_Storage Efficiency", paste0(filename2, "_eff.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 2, startCol = 5, units = "cm")
insertImage(wb, "Rain_norain_yr", paste0(filename2, "_rain_yr.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 2, startCol = 5, units = "cm")
insertImage(wb, "Rain_norain_mth", paste0(filename2, "_rain_mth.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 2, startCol = 5, units = "cm")
insertImage(wb, "Yield_Spill (Volume)", paste0(filename2, "_YS-vol.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 2, startCol = 5, units = "cm")
insertImage(wb, "Yield_Spill (Day)", paste0(filename2, "_YS-day.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 2, startCol = 5, units = "cm")
insertImage(wb, "Tank volume", paste0(filename2, "_tankvolper.jpg"),  
            width = 19.1, height = 11.58, 
            startRow = 2, startCol = 5, units = "cm")


# Save workbook to working directory
saveWorkbook(wb, file = paste0("TangkiNAHRIM_", filename2, "_summary.xlsx"), overwrite = TRUE)
