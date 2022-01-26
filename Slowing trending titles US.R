library.path <- .libPaths("C:/Users/Stephane/Documents/R/win-library/4.0")

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
setwd("C:\\Users\\Stephane\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv")

#Setting the date to read the correct previous week file
all_days <- c("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
pred_date <- Sys.Date() + 6 - match(weekdays(Sys.Date()), all_days) - 7

#Importing Data
Sales_US <- read.csv("Sales US.csv", header = T, stringsAsFactors = FALSE)
Sales_US$date <- as.Date(Sales_US$date)
Sales_US <- Sales_US[Sales_US$title != "",]


#Importing Data
forecasts <- read.csv( paste0("Forecast us - ",pred_date,".csv"), header = T, stringsAsFactors = FALSE)


for (i in 1:length(Sales_US$asin)){
  if (nchar(Sales_US$asin[i]) == 9){
    Sales_US$asin[i] <- paste0("0",Sales_US$asin[i])  
  }
  
}

for (i in 1:length(forecasts$asin)){
  if (nchar(forecasts$asin[i]) == 9){
    forecasts$asin[i] <- paste0("0",forecasts$asin[i])  
  }
  
}



#Creating the Data Frame with all sales data ------------------------------------------------------------------------------------
DF <- cbind.data.frame(Sales_US$date, Sales_US$asin, Sales_US$isbn, Sales_US$title,Sales_US$division ,Sales_US$pub_date,Sales_US$units ,stringsAsFactors = FALSE)
colnames(DF) <- c("date","asin","isbn","title","Division","Publication","units")

#Rearrange data with dates as columns
Q1 <- dcast(DF, asin + isbn + title + Division + Publication ~ date, value.var="units", fun.aggregate = sum)


DF <- merge(forecasts, Q1, by = "asin", all.x = TRUE)

DF$Error <- DF[,11] - DF[,ncol(DF)] 

DF2 <- DF[,c(1:5,11, (ncol(DF) - 1) ,  grep("Error", colnames(DF)) )]
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
addWorksheet(wb, sheetName="US")
writeData(wb, sheet="US", x=DF2)


#Creating borders function 
OutsideBorders <-
  function(wb_,
           sheet_,
           rows_,
           cols_,
           border_col = "black",
           border_thickness = "thick") {
    left_col = min(cols_)
    right_col = max(cols_)
    top_row = min(rows_)
    bottom_row = max(rows_)
    
    sub_rows <- list(c(bottom_row:top_row),
                     c(bottom_row:top_row),
                     top_row,
                     bottom_row)
    
    sub_cols <- list(left_col,
                     right_col,
                     c(left_col:right_col),
                     c(left_col:right_col))
    
    directions <- list("Left", "Right", "Top", "Bottom")
    
    mapply(function(r_, c_, d) {
      temp_style <- createStyle(border = d,
                                borderColour = border_col,
                                borderStyle = border_thickness)
      addStyle(
        wb_,
        sheet_,
        style = temp_style,
        rows = r_,
        cols = c_,
        gridExpand = TRUE,
        stack = TRUE
      )
      
    }, sub_rows, sub_cols, directions)
  }







#adding filters
addFilter(wb, "US", rows = 1, cols = 1:ncol(DF2))

#auto width for columns
width_vec <- suppressWarnings(apply(DF2, 2, function(x) max(nchar(as.character(x)) + 1, na.rm = TRUE)))
width_vec_header <- nchar(colnames(DF2))  + 4
max_vec_header <- pmax(width_vec, width_vec_header)
setColWidths(wb, "US",  cols = 1, widths = 13)
setColWidths(wb, "US",  cols = 2, widths = 15)
setColWidths(wb, "US",  cols = 3, widths = 50)
setColWidths(wb, "US",  cols = 6, widths = 8)

#Centering cells
centerStyle <- createStyle(halign = "center")
addStyle(wb, "US", style=centerStyle, rows = 2:nrow(DF2), cols = 6:ncol(DF2), 
         gridExpand = T, stack = TRUE)

leftStyle <- createStyle(halign = "left")
addStyle(wb, "US", style=leftStyle, rows = 2:nrow(DF2), cols = 1:5, 
         gridExpand = T, stack = TRUE)

# Adding borders
invisible(OutsideBorders(
  wb,
  sheet_ = "US",
  rows_ = 1:nrow(DF2),
  cols_ = 1:5
))

invisible(OutsideBorders(
  wb,
  sheet_ = "US",
  rows_ = 1:nrow(DF2),
  cols_ = 7:9
))

freezePane(
  wb,
  sheet = "US",
  firstActiveRow = 2
)

saveWorkbook(wb, paste0("Slowing trending titles US - ",pred_date +7,".xlsx"), overwrite = T) 



