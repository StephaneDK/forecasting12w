library.path <- .libPaths("C:/Users/steph/Documents/R/win-library/4.0")
source("C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\Code\\OutsideBorders.R")

cat("\nSlowing Trending Titles start\n")

oldw <- getOption("warn")
options(warn = -1)

suppressMessages({
  
  library(tidyverse, lib.loc = library.path)
  library(stringr, lib.loc = library.path)
  library(reshape2, lib.loc = library.path)
  library(ggthemes, lib.loc = library.path)
  library(gridExtra, lib.loc = library.path)
  library(forecast, lib.loc = library.path)
  library(aTSA, lib.loc = library.path)
  library(DescTools, lib.loc = library.path)
  library(plyr, lib.loc = library.path)
  library(EnvStats, lib.loc = library.path)
  library(qcc, lib.loc = library.path)
  library(openxlsx, lib.loc = library.path)
  library(magrittr, lib.loc = library.path)
})

options(warn = oldw)

options(scipen=999, digits = 3, error=function() { traceback(2); if(!interactive()) quit("no", status = 1, runLast = FALSE) } )


`%notin%` <- Negate(`%in%`)

#Setting the directory where all files will be used from for this project
setwd("C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv")

#Setting the date to read the correct previous week file
all_days <- c("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
pred_date <- Sys.Date() + 6 - match(weekdays(Sys.Date()), all_days) - 7

#Importing Data
Sales_UK <- read.csv("Sales uk.csv", header = T, stringsAsFactors = FALSE)
colnames(Sales_UK) <- tolower(colnames(Sales_UK) )
Sales_UK$date <- as.Date(Sales_UK$date)
Sales_UK <- Sales_UK[Sales_UK$title != "",]


#Importing Data
forecasts <- read.csv( paste0("Forecast uk - ",pred_date,".csv"), header = T, stringsAsFactors = FALSE)


for (i in 1:length(Sales_UK$asin)){
  if (nchar(Sales_UK$asin[i]) == 9){
    Sales_UK$asin[i] <- paste0("0",Sales_UK$asin[i])  
  }
  
}

for (i in 1:length(forecasts$asin)){
  if (nchar(forecasts$asin[i]) == 9){
    forecasts$asin[i] <- paste0("0",forecasts$asin[i])  
  }
  
}



#Creating the Data Frame with all sales data ------------------------------------------------------------------------------------
DF <- cbind.data.frame(Sales_UK$date, Sales_UK$asin, Sales_UK$isbn, Sales_UK$title,Sales_UK$division ,Sales_UK$pub_date,Sales_UK$units ,stringsAsFactors = FALSE)
colnames(DF) <- c("date","asin","isbn","title","Division","Publication","units")

#Rearrange data with dates as columns
Q1 <- dcast(DF, asin + isbn + title + Division + Publication ~ date, value.var="units", fun.aggregate = sum)





DF <- merge(forecasts, Q1, by = "isbn", all.x = TRUE)


DF$Error <-   DF[,c(11)] - DF[,c(ncol(DF) )]

DF2 <- DF[,c(1:5,(11), (ncol(DF) - 1) ,  grep("Error", colnames(DF)) )]
colnames(DF2) <- c("asin", "isbn", "Title", "Division", "Publication", "Forecast", "Actual",  "Error")


DF2$Actual[is.na(DF2$Actual)] <- 0
DF2$Error[is.na(DF2$Error)] <- 0



#-----------------------------------------------------------------------------------------------------------#
#                                     Saving output formatting                                              #
#-----------------------------------------------------------------------------------------------------------#

# Adding empty columns
DF2 <- add_column(DF2, new_col = NA, .after = 5)


colnames(DF2)[6] <- ""



#Creating a workbook
wb <- createWorkbook()
options("openxlsx.numFmt" = "0")
addWorksheet(wb, sheetName="UK")
writeData(wb, sheet="UK", x=DF2)


#adding filters
addFilter(wb, "UK", rows = 1, cols = 1:ncol(DF2))

#auto width for columns
width_vec <- suppressWarnings(apply(DF2, 2, function(x) max(nchar(as.character(x)) + 1, na.rm = TRUE)))
width_vec_header <- nchar(colnames(DF2))  + 4
max_vec_header <- pmax(width_vec, width_vec_header)
setColWidths(wb, "UK",  cols = 1, widths = 16)
setColWidths(wb, "UK",  cols = 2, widths = 12)
setColWidths(wb, "UK",  cols = 3, widths = 65)
setColWidths(wb, "UK",  cols = 4, widths = 20)
setColWidths(wb, "UK",  cols = 5, widths = 13)
setColWidths(wb, "UK",  cols = 6, widths = 8)

#Centering cells
centerStyle <- createStyle(halign = "center")
addStyle(wb, "UK", style=centerStyle, rows = 2:nrow(DF2), cols = 6:ncol(DF2), 
         gridExpand = T, stack = TRUE)

leftStyle <- createStyle(halign = "left")
addStyle(wb, "UK", style=leftStyle, rows = 2:nrow(DF2), cols = 1:5, 
         gridExpand = T, stack = TRUE)

# Adding borders
invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(DF2),
  cols_ = 1:5
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(DF2),
  cols_ = 7:9
))

freezePane(
  wb,
  sheet = "UK",
  firstActiveRow = 2
)

saveWorkbook(wb, paste0("Slowing trending titles UK - ",pred_date +7,".xlsx"), overwrite = T) 


cat("Slowing Trending Titles end\n")


