setwd("L:/Prices/AMR/Livestock/MONTHLY/Machine readable")
library(readxl)
library(tidyr)
library(stringr)
library(data.table)
library(dplyr)

##Read in all sheets
read_excel_allsheets <- function(filename, skip = 7, tibble = T) {
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X, skip = skip, col_names = F, na = c("n/a", "*na")))
  if(!tibble) x <- lapply(x, as.data.frame)
  names(x) <- sheets
  x
}

##Convert finished livestock prices####

##Read in all sheets
finished_all_sheets_list <- read_excel_allsheets("livestkmth Finished.xls", skip = 8)
finished_all_sheets_list <- Map(cbind, finished_all_sheets_list)

##Convert to one long dataframe
livestock_finished_all_years <- rbindlist(finished_all_sheets_list, fill=T)

##Tidy dataframe
livestock_finished <- livestock_finished_all_years[,c(1:3,5,7)]

##Rename columns
colnames(livestock_finished) <- c("Year", "Month", "clean_cattle", "sheep", "all_pigs")

##Gather into long form 
livestock_finished <- livestock_finished %>% 
  gather(Livestock, Price, -c(Year, Month)) 

##Convert price to numerics 
livestock_finished$Price <- round(as.numeric(livestock_finished$Price), 2)

##Add units and remove all non-numeric values
livestock_finished <- livestock_finished %>% na.omit() %>% 
  mutate(Units = case_when(Livestock == "clean_cattle" ~ "p/kg liveweight",
                           TRUE ~ "p/kg deadweight"))
#Arrange by year
livestock_finished <- livestock_finished %>% arrange(desc(Year))

##Write final csv
write.csv(livestock_finished, "livestock_finished_csv.csv", row.names = F)

##Convert stores livestock prices####

##Read in all sheets
stores_all_sheets_list <- read_excel_allsheets("livestkmth Stores.xls", skip = 8)
stores_all_sheets_list <- Map(cbind, stores_all_sheets_list, Year = names(stores_all_sheets_list))

##Convert to one long dataframe
livestock_stores_all_years <- rbindlist(stores_all_sheets_list, fill=T)

##Select relevant columns
livestock_stores <- livestock_stores_all_years[,-c(28,29)]

##Rename first three columns
colnames(livestock_stores)[1:3] <- c("Livestock", "Age", "Subtype")

##Add in type column: type is equal to livestock in heading rows only (so na subtype)
livestock_stores <- livestock_stores %>% mutate(Type = case_when(is.na(Subtype) ~ Livestock))

##Fill down metadata
livestock_stores <- livestock_stores %>% fill(Livestock) %>% fill(Type) %>% fill(Age)

##Remove na values for subtype
livestock_stores <- livestock_stores[!is.na(livestock_stores$Subtype),]

##Clean age column
livestock_stores$Age <- gsub("\\(.*", "", livestock_stores$Age)
livestock_stores$Age <- str_trim(livestock_stores$Age)
livestock_stores$Age <- tolower(livestock_stores$Age)
livestock_stores$Age <- gsub(" ", "_", livestock_stores$Age)
livestock_stores$Age <- gsub("*others*", "", livestock_stores$Age)
livestock_stores$Age <- gsub("*cross*", "", livestock_stores$Age)
livestock_stores$Age <- gsub("*friesian*", "", livestock_stores$Age)

##Clean year column
livestock_stores$Year <- gsub(" .*", "", livestock_stores$Year)

##Clean type column
livestock_stores$Type <- gsub("\\(.*", "", livestock_stores$Type)
livestock_stores$Type <- str_trim(livestock_stores$Type)
livestock_stores$Type <- tolower(livestock_stores$Type)
livestock_stores$Type <- gsub(" ", "_", livestock_stores$Type)

##Clean breed column
livestock_stores$Livestock <- str_trim(livestock_stores$Livestock)
livestock_stores$Livestock <- tolower(livestock_stores$Livestock)
livestock_stores$Livestock <- gsub("*store.*", "", livestock_stores$Livestock)
livestock_stores$Livestock <- gsub(" ", "_", livestock_stores$Livestock)

##Clean subtype column 
livestock_stores$Subtype <- str_trim(livestock_stores$Subtype)
livestock_stores$Subtype <- tolower(livestock_stores$Subtype)
livestock_stores$Subtype <- gsub("[[:punct:]]", "", livestock_stores$Subtype)
livestock_stores$Subtype <- gsub(" ", "_", livestock_stores$Subtype)

##Add names to clean columns
colnames(livestock_stores) <- c("Breed", "Age", "Subtype", "Jan_count","Jan_price", 
                                "Feb_count", "Feb_price", "Mar_count", "Mar_price", 
                                "Apr_count", "Apr_price", "May_count", "May_price",
                                "Jun_count", "Jun_price", "Jul_count", "Jul_price", 
                                "Aug_count", "Aug_price", "Sep_count", "Sep_price",
                                "Oct_count", "Oct_price", "Nov_count", "Nov_price",
                                "Dec_count", "Dec_price", "Year", "Type")



##Order columns and gather into long form 
livestock_stores <- livestock_stores %>% 
  select(Type, Breed, Age, Subtype, Year, everything()) %>%
  gather(Month, Value, -c(Type:Year)) 

##Split month/data type data and add units
livestock_stores <- livestock_stores %>%
  separate(Month, c("Month", "Data_type"), "_") %>%
  mutate(Units = case_when(Data_type == "price" ~ "?/head",
                           Data_type == "count" ~ "head"))

##Remove zero prices and counts
livestock_stores <- livestock_stores %>% filter(Value != 0)

#Arrange by year
livestock_stores <- livestock_stores %>% arrange(desc(Year))

##Write final csv
write.csv(livestock_stores, "livestock_stores_csv.csv", row.names = F)


