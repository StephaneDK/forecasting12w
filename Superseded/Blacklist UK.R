library.path <- .libPaths("C:/Users/steph/Documents/R/win-library/4.0")
cat("Forecast start\n")

suppressMessages(library(tidyverse, lib.loc = library.path))
suppressMessages(library(stringr, lib.loc = library.path))
suppressMessages(library(reshape2, lib.loc = library.path))
suppressMessages(library(ggthemes, lib.loc = library.path))
suppressMessages(library(gridExtra, lib.loc = library.path))
suppressMessages(library(forecast, lib.loc = library.path))
suppressMessages(library(aTSA, lib.loc = library.path))
suppressMessages(library(DescTools, lib.loc = library.path))
suppressMessages(library(plyr, lib.loc = library.path))
suppressMessages(library(EnvStats, lib.loc = library.path))
suppressMessages(library(qcc, lib.loc = library.path))
suppressMessages(library(openxlsx, lib.loc = library.path))

options(scipen=999, digits = 3, error=function() { traceback(2); if(!interactive()) quit("no", status = 1, runLast = FALSE) } )

all_days <- c("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")

`%notin%` <- Negate(`%in%`)
current_quarter <- "Q3"

#Setting the directory where all files will be used from for this project
setwd("C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv")

#Importing scrape data (product availability)
Scrape_UK <- read.csv("Scrape info uk.csv", header = T, stringsAsFactors = FALSE)
colnames(Scrape_UK) <- c("asin", "AMZ Availability")

for (i in 1:length(Scrape_UK$asin)){
  if (nchar(Scrape_UK$asin[i]) == 9){
    Scrape_UK$asin[i] <- paste0("0",Scrape_UK$asin[i])  
  }
  
}

AMZ_Stock <- read.csv("AMZ Stock status.csv", header = T, stringsAsFactors = FALSE)

AMZ_Stock <- AMZ_Stock %>% 
  subset(country == "uk")  %>% 
  select(isbn, days_oos, days_lowstock)


Blacklist <- readRDS("pred_df_holt_damp_beta_UK.rds")

Blacklist <- suppressWarnings(merge(Blacklist, AMZ_Stock, by = "isbn", all.x = T))

Blacklist <- Blacklist  %>%
             mutate(days_oos = as.numeric(days_oos),
                    days_lowstock = as.numeric(days_lowstock),
                    l4w  = rowSums(.[7:10]) )


#-----------------------------------------------------------------------------------------------------------------------
#                                 Save 1
#-----------------------------------------------------------------------------------------------------------------------

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





save1 <- Blacklist

save1 <- save1[,c(1:6,24:41)]

# Adds after the second column
save1 <- add_column(save1, new_col = NA, .after = 21)

colnames(save1)[6] <- ""
colnames(save1)[12] <- ""
colnames(save1)[17] <- ""
colnames(save1)[22] <- ""


#Highlighting columns with inventory issues
wb <- createWorkbook()
addWorksheet(wb, sheetName="UK")
writeData(wb, sheet="UK", x=save1)


#adding filters
addFilter(wb, "UK", rows = 1, cols = 1:ncol(save1))

#auto width for columns
width_vec <- suppressWarnings(apply(save1, 2, function(x) max(nchar(as.character(x)) + 1, na.rm = TRUE)))
width_vec_header <- nchar(colnames(save1))  + 3
max_vec_header <- pmax(width_vec, width_vec_header)
setColWidths(wb, "UK", cols = 1:ncol(save1), widths = max_vec_header )
setColWidths(wb, "UK",  cols = 1, widths = 13)
setColWidths(wb, "UK",  cols = 2, widths = 15)
setColWidths(wb, "UK",  cols = 3, widths = 52)
setColWidths(wb, "UK",  cols = 6, widths = 10)
setColWidths(wb, "UK",  cols = 10, widths = 10)
setColWidths(wb, "UK",  cols = 11, widths = 14)
setColWidths(wb, "UK",  cols = 12, widths = 10)
setColWidths(wb, "UK",  cols = 17, widths = 10)
setColWidths(wb, "UK",  cols = 22, widths = 10)



#Centering cells
centerStyle <- createStyle(halign = "center")
addStyle(wb, "UK", style=centerStyle, rows = 2:nrow(save1), cols = 7:ncol(save1), 
         gridExpand = T, stack = TRUE)

# Adding borders
invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(save1)+1,
  cols_ = 1:5
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(save1)+1,
  cols_ = 7:11
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(save1)+1,
  cols_ = 13:16
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(save1)+1,
  cols_ = 18:21
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(save1)+1,
  cols_ = 23:25
))

freezePane(
  wb,
  sheet = "UK",
  firstActiveRow = 2,
  firstActiveCol = 6
)


pred_date <- Sys.Date() + 6 - match(weekdays(Sys.Date()), all_days)

saveWorkbook(wb, paste0("OOS Review uk - ",pred_date,".xlsx"), overwrite = T ) 




#-----------------------------------------------------------------------------------------------------------------------
#                                 Blacklist rules
#-----------------------------------------------------------------------------------------------------------------------




#Subset blacklisted titles
Blacklist <- Blacklist[ (Blacklist$`AA Status` %in% c("Super","Support") & Blacklist$WOH != "12+") | 
                        (Blacklist$`AA Status` %in% c("Super","Support") & 
                         Blacklist$days_oos >= 2 & 
                         Blacklist$inventory <= 150 &
                         (Blacklist$Reprint.Date >= "2022-01-01" | is.na(Blacklist$Reprint.Date) )  ),]

Blacklist <- Blacklist[!is.na(Blacklist$title),]



#-----------------------------------------------------------------------------------------------------------------------
#                                 Save 1
#-----------------------------------------------------------------------------------------------------------------------


Blacklist <- Blacklist[,c(1:6,24,27:37,41,38:40)]

colnames(Blacklist)[7] <- "12W_Forecast"

# Adds after the second column
Blacklist <- add_column(Blacklist, new_col = NA, .after = 20)
Blacklist <- add_column(Blacklist, new_col = NA, .after = 18)

colnames(Blacklist)[6] <- ""
colnames(Blacklist)[10] <- ""
colnames(Blacklist)[15] <- ""
colnames(Blacklist)[19] <- ""
colnames(Blacklist)[22] <- ""


#Highlighting columns with inventory issues
wb <- createWorkbook()
addWorksheet(wb, sheetName="UK")
writeData(wb, sheet="UK", x=Blacklist)


#adding filters
addFilter(wb, "UK", rows = 1, cols = 1:ncol(Blacklist))

#auto width for columns
width_vec <- suppressWarnings(apply(Blacklist, 2, function(x) max(nchar(as.character(x)) + 1, na.rm = TRUE)))
width_vec_header <- nchar(colnames(Blacklist))  + 3
max_vec_header <- pmax(width_vec, width_vec_header)
setColWidths(wb, "UK", cols = 1:ncol(Blacklist), widths = max_vec_header )
setColWidths(wb, "UK",  cols = 1, widths = 13)
setColWidths(wb, "UK",  cols = 2, widths = 15)
setColWidths(wb, "UK",  cols = 3, widths = 52)
setColWidths(wb, "UK",  cols = 6, widths = 10)
setColWidths(wb, "UK",  cols = 10, widths = 10)
setColWidths(wb, "UK",  cols = 9, widths = 14)
setColWidths(wb, "UK",  cols = 15, widths = 10)
setColWidths(wb, "UK",  cols = 19, widths = 10)
setColWidths(wb, "UK",  cols = 22, widths = 10)


# Highlighting rows in red
red_style <- createStyle(fgFill="#FF0000")
x <- which( Blacklist$WOH %in% c("1","2") & (Blacklist$Reprint.Date >= "2021-12-25" | is.na(Blacklist$Reprint.Date) ) )
addStyle(wb, sheet="UK", style=red_style, rows=x+1, cols=c(1:ncol(Blacklist)), 
         gridExpand=TRUE, stack = TRUE) 

# Highlighting rows in orange
orange_style <- createStyle(fgFill="#FFA500")
x <- which( Blacklist$WOH %in% c("3","4") & (Blacklist$Reprint.Date >= "2021-12-25" | is.na(Blacklist$Reprint.Date) ) )
addStyle(wb, sheet="UK", style=orange_style, rows=x+1, cols=c(1:ncol(Blacklist)), 
         gridExpand=TRUE, stack = TRUE) 



#Centering cells
centerStyle <- createStyle(halign = "center")
addStyle(wb, "UK", style=centerStyle, rows = 2:nrow(Blacklist), cols = 7:ncol(Blacklist), 
         gridExpand = T, stack = TRUE)

# Adding borders
invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(Blacklist)+1,
  cols_ = 1:5
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(Blacklist)+1,
  cols_ = 7:9
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(Blacklist)+1,
  cols_ = 11:14
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(Blacklist)+1,
  cols_ = 16:18
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(Blacklist)+1,
  cols_ = 20:21
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(Blacklist)+1,
  cols_ = 23:24
))


freezePane(
  wb,
  sheet = "UK",
  firstActiveRow = 2,
  firstActiveCol = 6
)


pred_date <- Sys.Date() + 6 - match(weekdays(Sys.Date()), all_days)

saveWorkbook(wb, paste0("Blacklist uk - ",pred_date,".xlsx"), overwrite = T ) 




