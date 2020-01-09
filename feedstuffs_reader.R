setwd("L:/Prices/AMR/Feedingstuffs/Machine readable")
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

#Remove columns that are all na
not_all_na <- function(x) any(!is.na(x))

##Add column to with year name to each row
all_sheets_list <- read_excel_allsheets("L:/Prices/AMR/Feedingstuffs/NewStraight Feedingstuffs.xlsx")
all_sheets_list <- Map(cbind, all_sheets_list, Year = names(all_sheets_list))

##Convert to one long dataframe
feedstuffs_all_years <- rbindlist(all_sheets_list, fill=T)

##Tidy dataframe
feedstuffs <- feedstuffs_all_years %>% select_if(not_all_na)
##Rename columns
colnames(feedstuffs) <- c("Feedstuff", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Year")
##Gather into long form
feedstuffs <- feedstuffs %>% gather(Month, Price, -c(Feedstuff, Year))
##Convert price to numerics and remove all non-numeric values
feedstuffs$Price <- round(as.numeric(feedstuffs$Price), 2)
feedstuffs <- feedstuffs %>% na.omit() %>% mutate(Units = "?/ton")

##Tidy up feed names
feedstuffs$Feedstuff <- gsub("[[:punct:]].*", "\\1", feedstuffs$Feedstuff)
feedstuffs$Feedstuff <- gsub("supaflow", "", feedstuffs$Feedstuff)
feedstuffs$Feedstuff <- gsub("EU", "", feedstuffs$Feedstuff)
feedstuffs$Feedstuff <- gsub("[0-9]", "", feedstuffs$Feedstuff)
feedstuffs$Feedstuff <- tolower(feedstuffs$Feedstuff)
feedstuffs$Feedstuff <- str_trim(feedstuffs$Feedstuff)
feedstuffs$Feedstuff <- gsub(" ", "_", feedstuffs$Feedstuff)

##Write final csv
write.csv(feedstuffs, "feedstuffs_csv.csv", row.names = F)





