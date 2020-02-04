setwd("L:/Prices/API/API 2015/Machine readable")

if (!require("pacman")) install.packages("pacman")
library(pacman)
pacman::p_load(dplyr, readxl, data.table, stringr, tidyr, janitor, lubridate)

library(readxl)
library(tidyr)
library(stringr)
library(data.table)
library(dplyr)
library(janitor)
library(lubridate)

##Read in all sheets
read_excel_allsheets <- function(filename, skip = 7, tibble = T) {
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X, skip = skip, col_names = T, na = c("n/a", "*na")))
  if(!tibble) x <- lapply(x, as.data.frame)
  names(x) <- sheets
  x
}

##Read in all sheets
finished_all_sheets_list <- read_excel_allsheets("L:/Prices/API/API 2015/Stats Notice/API datatset_working.xlsx", skip = 5)
finished_all_sheets_list <- Map(cbind, finished_all_sheets_list)

##Remove annual sheets
finished_all_sheets_list <- finished_all_sheets_list[-c(grep("annual", names(finished_all_sheets_list)))]

##Convert to one long dataframe
API_all_years <- rbindlist(finished_all_sheets_list, fill=T)

##Tidy dataframe
API_years <- API_all_years[,-c(2,3)]

##Rename columns
colnames(API_years)[1] <- "Category"
colnames(API_years) <- gsub("[(].*", "", colnames(API_years))

##Remove unnamed categories
API_years <- API_years %>% 
  filter(!is.na(Category)) %>%
  gather(Date, Index, -c(Category)) 

##Convert date to date format; Date_1 converts those in excel numeric format,
#Date_2 converts those in text format
API_years$Date_1 <- excel_numeric_to_date(as.numeric(API_years$Date))
API_years$Date_2 <- as.Date(paste0("01-", API_years$Date), format = "%d-%b-%y")

##If function which selects date from either Date column 1 or 2
API_years <- API_years %>% 
  mutate(Date = case_when(!is.na(Date_2) ~ Date_2,
                          TRUE ~ Date_1)) %>%
  select(-c(Date_1, Date_2))

##Round indexes to 2dp
API_years$Index <- round(as.numeric(API_years$Index),2)

##Convert dates to years and months
API_years <- API_years %>% 
  mutate(Year = year(Date), Month = format(Date, "%b")) %>%
  select(-Date) %>%
  ##Remove any NAs
  na.omit()

##Tidy descriptors; 
#all lower case
API_years$Category <- tolower(API_years$Category)
#Remove extra spaces and punctuation
API_years$Category <- gsub("  ", " ", API_years$Category)
API_years$Category <- gsub("[[:punct:]]", "", API_years$Category)
API_years$Category <- str_squish(API_years$Category)
#Convert all spaces to underscores
API_years$Category <- gsub(" ", "_", API_years$Category)

#Arrange by year
API_years <- API_years %>% arrange(desc(Year))

##Write final csv
write.csv(API_years, "API_machine_readable.csv", row.names = F)

