library.path <- .libPaths("C:/Users/steph/Documents/R/win-library/4.0")
source("C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\Code\\OutsideBorders.R")
source("C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\Code\\ChristmasAdjustment.R")
source("C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\Code\\TopTitlesAdjustment.R")

cat("\nForecast start\n")

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

options(scipen=999, digits = 3 )
all_days <- c("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")

`%notin%` <- Negate(`%in%`)

convert.brackets <- function(x){
  if(grepl("\\(.*\\)", x)){
    paste0("-", gsub("\\(|\\)", "", x))
  } else {
    x
  }
}

current_quarter <- c("Q2")

#Setting the directory where all files will be used from for this project
setwd("C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv")

#Importing and clean Sales Data
Sales_US <- read.csv("Sales US Bookscan.csv", header = T, stringsAsFactors = FALSE)
Sales_US <- Sales_US %>% 
  rename_all(tolower)  %>% 
  subset(title != "") %>% 
  mutate(date = as.Date(date),
         asin = case_when(nchar(asin) == 9 ~ paste0("0",asin),
                          TRUE ~ asin)
  ) %>%
  select(date, asin, isbn, title, division, pub_date, units) %>%
  set_colnames(c("date","asin","isbn","title","Division","Publication","units")) %>%
  dcast(asin + isbn + title + Division + Publication ~ date, value.var="units", fun.aggregate = sum) %>%
  subset( rowSums( .[,(ncol(.)-3): ncol(.)]) >= 5 ) %>%
  arrange(desc(.[ncol(.)])) %>%
  mutate_all(~replace(., is.na(.), 0))

Q1 <- Sales_US

#Import AA Data
AA_status <-  read.csv("AA Status us.csv", header = T, stringsAsFactors = FALSE)
colnames(AA_status) <- c("isbn", "AA_status")

Print_status <-  read.csv("Print Status us.csv", header = T, stringsAsFactors = FALSE)
colnames(Print_status) <- c(  "isbn", "Print_status")


#Importing reprint dates & quantity
Reprint2 <- read.csv("Reprint.csv", header = T, stringsAsFactors = FALSE) 
#sort(colnames(Reprint2))
Reprint2 <- Reprint2 %>% 
  subset(Print.Instruction.Impression %in% c("DK US", "DK US (Alpha/Brady)", "DK US (DK RG SS)", "DK US (DK RG TR)", 
                                             "DK US (DK RG TR) USX", "DK US (USX)", "DK US REBEL GIRLS") ) %>% 
  mutate(Schedule.Date = as.Date(Schedule.Date, format = "%d/%m/%Y")) %>% 
  select( Print.Instruction.Edition, Schedule.Date,  Print.Instruction.Quantity) %>% 
  group_by(Print.Instruction.Edition) %>%
  filter(Schedule.Date == min(Schedule.Date) )  %>%
  top_n(1, Print.Instruction.Quantity) %>%
  set_colnames(c( "ISBN", "Reprint.Date", "Reprint Qty" ))



#Importing Inventory
Inventory <- read.csv("Inventory us.csv", header = T, stringsAsFactors = FALSE) 

Inventory <- Inventory %>% 
  select( ASIN, MATERIAL, RETAIL_INV, TBS_INV, AMZ_OPEN_ORDERS) %>% 
  set_colnames(c("asin" ,"ISBN", "Retail inv", "TBS inv", "AMZ Open Orders" )) %>%
  merge( Reprint2, by = "ISBN", all.x = T) %>%
  mutate(title = "NA") %>%
  select(asin, title, `Retail inv`, `TBS inv`, Reprint.Date, 
         `AMZ Open Orders`, `Reprint Qty`) %>%
  mutate(Reprint.Date = Reprint.Date + 6 - match(weekdays(Reprint.Date), all_days),
         `TBS inv` = ifelse(is.na(`TBS inv`), 0, `TBS inv`),
         `Retail inv` =  ifelse(is.na(`Retail inv`), 0, `Retail inv`),
         `Reprint Qty` = replace(`Reprint Qty`, is.na(`Reprint Qty`), NaN) )

Reprint <- Inventory


#Import Seasonal titles ---------------------------------------------------------------------------------------------------

Q_iso <- readRDS("Q_iso_US.csv")


#Adding seasonal titles manually (if missing)
Q_iso <- rbind.data.frame(Q_iso, c("0241459001","9780241459003","Baby's First Easter","Q1")  )
Q_iso <- rbind.data.frame(Q_iso, c("1465489665","9781465489661","Baby's First St. Patrick's Day","Q1")  )
Q_iso <- rbind.data.frame(Q_iso, c("1465494855","9781465494856","Ultimate Sticker Book Passover","Q1")  )
Q_iso <- rbind.data.frame(Q_iso, c("0744069866","9780744069860","The Sleepy Bunny","Q1")  )
Q_iso <- rbind.data.frame(Q_iso, c("0744026598","9780744026597","Baby's First Ramadan","Q1")  )

Q_iso <- rbind.data.frame(Q_iso, c("0744069866","9780744069860","The Sleepy Bunny","Q2")  )
Q_iso <- rbind.data.frame(Q_iso, c("0241459001","9780241459003","Baby's First Easter","Q2")  )
Q_iso <- rbind.data.frame(Q_iso, c("1465494855","9781465494856","Ultimate Sticker Book Passover","Q2")  )
Q_iso <- rbind.data.frame(Q_iso, c("0744026598","9780744026597","Baby's First Ramadan","Q2")  )

Q_iso <- rbind.data.frame(Q_iso, c("146546235X","9781465462350","Baby Touch and Feel: Halloween","Q4")  )
Q_iso <- rbind.data.frame(Q_iso, c("1465452761","9781465452764","Pop-Up Peekaboo! Pumpkin","Q4")  )
Q_iso <- rbind.data.frame(Q_iso, c("0744033837","9780744033830","The Happy Pumpkin","Q4")  )


#Creating computational statistics ---------------------------------------------------------------------------------------------------

#Time Vectors
time_vec <- c(1:4)
time_vec_future <- c(5:16)


#-----------------------------------------------------------------------------------------------------------------------
#                                 Holt model prediction
#-----------------------------------------------------------------------------------------------------------------------

a <- matrix(, nrow = nrow(Q1), ncol = 2)
b <- matrix(, nrow = nrow(Q1), ncol = length(time_vec_future))

for (i in 1:nrow(Q1)){
  
  #Actual Holt function from forecast package
  a[i,] <- (Holt(as.numeric(Q1[i,(ncol(Q1) - 3):ncol(Q1)]), lead = 2, plot = FALSE, type = "additive")$pred)
  b[i,] <- (Holt(as.numeric(Q1[i,(ncol(Q1) - 3):ncol(Q1)]), lead = length(time_vec_future),phi = 0.9, damped = TRUE, plot = FALSE, type = "additive")$pred)
  
}

#Storing results matrix into a data frame
pred_df_holt_damp_beta <- cbind.data.frame(Q1[,c(1:5,(ncol(Q1) - 3):ncol(Q1))],a, b[,3:ncol(b)])
  

#Renaming columns

for (i in 1:ncol(pred_df_holt_damp_beta)){
 
  if (i > 9){

    colnames(pred_df_holt_damp_beta)[i] <-  as.character( as.Date(colnames(pred_df_holt_damp_beta)[i-1]) + 7 )
  }
}


#Adjusting for predictions with big negative slopes
for (i in 1:nrow(pred_df_holt_damp_beta)){
  
  if (round(as.numeric(lm(as.numeric(pred_df_holt_damp_beta[i,c(6:9)]) ~ c(1:4))$coefficients[2]),1) <= -5){
    
    for (j in 10:21){
      pred_df_holt_damp_beta[i,j] <- max(pred_df_holt_damp_beta[i,j], 
                                         rowMeans(pred_df_holt_damp_beta[i,6:9])/2.2)
      
      pred_df_holt_damp_beta[i,j] <- min(pred_df_holt_damp_beta[i,j], rowMeans(pred_df_holt_damp_beta[i,8:9]) )
      
      
    }
    
  }
  
}

#-----------------------------------------------------------------------------------------------------------------------
#                                 Christmas period adjustment
#-----------------------------------------------------------------------------------------------------------------------

pred_df_holt_damp_beta <- ChristmasAdjustment(pred_df_holt_damp_beta, Q1)

#-----------------------------------------------------------------------------------------------------------------------
#                                 Seasonal titles adjustment
#-----------------------------------------------------------------------------------------------------------------------

#Create 2020 only data frame
prev_year <- Q1[ Q1$asin %in% Q_iso$asin, ]
prev_year <- prev_year[,c(1,3,grep("2021-01-02", colnames(Q1)): grep("2022-09-17", colnames(Q1))) ]
prev_year[prev_year <= 0] <- 1


#Create empty data frame to store seasonal percentage changes
Seas_adjQ <- data.frame(matrix(ncol = ncol(prev_year) , nrow = nrow(prev_year)))
colnames(Seas_adjQ) <- colnames(prev_year)

Seas_adjQ$asin <- prev_year$asin
Seas_adjQ$title <- prev_year$title



#Compute seasonal percentage changes using last year data
for (i in 1:nrow(Seas_adjQ)){
  for ( j in 1:ncol(Seas_adjQ)){
    if (j == 3){
      Seas_adjQ[i,j] <- 1
        
    } else if (j >3) {
      Seas_adjQ[i,j] <- prev_year[i,j] / prev_year[i,j-1]
        
    }
    
  }
  
}

#Clean and readjust
Seas_adjQ[is.na(Seas_adjQ)] <- 1
Seas_adjQ[Seas_adjQ == 0] <- 1
Seas_adjQ[Seas_adjQ == Inf] <- 1

#Re-organise

#Manual seasonal adjustment
Seas_adjQ[Seas_adjQ$asin == "1465484019",3:16]  <- c( 1, 1.6, 1, 1.5, 1.2, 1.2, 1.5, 0.27, 0.44, 0.75, 0.66, 1, 1, 1)
Seas_adjQ[Seas_adjQ$asin == "146546526X",3:16]  <- c( 1, 4.2, 1.3, 1.5, 1.1, 1.5, 1.5, 0.1, 0.6, 0.66, 0.91, 1.2, 1, 0.33)
Seas_adjQ[Seas_adjQ$asin == "1465457631",3:16] <-  c(1, 2.4, 1.9, 1.9, 1.7, 1.5, 1.5, 0.1, 0.54, 0.6, 0.62, 1.4, 1, 1)

#Rounding to 3 decimal places
Seas_adjQ[,3:ncol(Seas_adjQ)] <- round(Seas_adjQ[,3:ncol(Seas_adjQ)],3)



#Keep only current seasonal adjustment data
Seas_adjQ <- Seas_adjQ[ Seas_adjQ$asin %in% subset(Q_iso, Q_iso$Season %in% current_quarter)$asin ,]
Seas_adjQ <- arrange(Seas_adjQ, desc(Seas_adjQ$asin))



pred_df_holt_damp_beta <- arrange(pred_df_holt_damp_beta, desc(pred_df_holt_damp_beta$asin)) 

#Seasonal adjustment loop
for (i in 1:nrow(pred_df_holt_damp_beta)){
 
  
  if (pred_df_holt_damp_beta$asin[i] %in% Seas_adjQ$asin ){
    
    for (j in 10:21){
      
      pred_df_holt_damp_beta[i,j] <- pred_df_holt_damp_beta[i,j-1]*Seas_adjQ[ Seas_adjQ$asin == pred_df_holt_damp_beta$asin[i], 
                                                                              as.character(as.Date(colnames(pred_df_holt_damp_beta)[j]) - 364 ) ]
      
    }
  }
}


#-----------------------------------------------------------------------------------------------------------------------
#                                 Category Adjustment 
#-----------------------------------------------------------------------------------------------------------------------


#Last 7 weeks of travel books are increased by 20%
for ( i in 1:nrow(pred_df_holt_damp_beta)){
  if (pred_df_holt_damp_beta$Division[i] == "Travel"){
    pred_df_holt_damp_beta[i,10:21] <-  pred_df_holt_damp_beta[i,10:21]*1.2 
  }
}

#Manual adjustment for books with PR campaigns or sudden sales spikes (not used)



#-----------------------------------------------------------------------------------------------------------------------
#                                 Title level adjustment for top 55 titles
#-----------------------------------------------------------------------------------------------------------------------


#pred_df_holt_damp_beta <- TopTitlesAdjustment(pred_df_holt_damp_beta, Q1)



#Replacing negative predictions with 0 and rounding all predictions to closest number
for ( i in 1:nrow(pred_df_holt_damp_beta)){

  for ( j in 10:21){


    pred_df_holt_damp_beta[i,j] <- round(pred_df_holt_damp_beta[i,j], digits = 0)

    if (pred_df_holt_damp_beta[i,j]<0){

      pred_df_holt_damp_beta[i,j] <- 0

    }
  }
}

pred_df_holt_damp_beta$Total_12w <- rowSums(pred_df_holt_damp_beta[,c(10:21) ])

#-----------------------------------------------------------------------------------------------------------------------
#                                 Inventory Adjustment
#-----------------------------------------------------------------------------------------------------------------------

#Merge reprint & Sales Data Frames
DF <- merge(pred_df_holt_damp_beta, Reprint, by = "asin", all.x = TRUE)
DF[is.na(DF$`AMZ Open Orders`),27] <- 0

#Replace negative values of inventory in TBS by 0 (not meaningful)
for ( i in 1:nrow(DF)){
  
  if (DF$`TBS inv`[i] < 0 & !is.na(DF$`TBS inv`[i])){
    DF$`TBS inv`[i] <- 0
  }
  
  if (DF$`Retail inv`[i] < 0 & !is.na(DF$`Retail inv`[i])){
    DF$`Retail inv`[i] <- 0
  }
  
}

#One variable for amazon and tbs inventory
DF$inventory <- DF$`Retail inv` + DF$`TBS inv` #+ DF$`AMZ Open Orders`
DF$Inv.issue <- 0
DF$WOH <- "12+"

DF$Publication <- as.Date(DF$Publication)


for (i in 1:nrow(DF)){
  
  #If the sum of projected sales is lower than available inventory, already published, and non null
  if ( !is.na(DF$inventory[i]) & DF$Publication[i] < Sys.Date() 
     & !is.na(DF$Publication[i]) & DF$inventory[i]  < as.numeric(rowSums(DF[i,10:21])) ){ 
    
      a = 1
      b = 1
      temp_value <- DF$inventory[i]
      
      for ( j in 10:21){
        
        
        if ( as.Date(colnames(DF[j])) < DF$Reprint.Date[i] | is.na(DF$Reprint.Date[i]) ){
          
          
          if ( DF$inventory[i] - DF[i,j] > 0 ){
            
            DF$inventory[i] <- DF$inventory[i] - DF[i,j]
            temp_value <-  DF$inventory[i]
            b = b + 1
            
          } else if ( DF$inventory[i] - DF[i,j]  <= 0){
            
            if (a == 1){
              DF[i,j] <- temp_value
              DF$WOH[i] <- b - 1
              DF$inventory[i] <- 0
              a = a + 1
            } else {
              DF[i,j] <- 0
              DF$WOH[i] <- b }
            
            
            if ( rowSums(DF[i,6:9]) > 30 ){
              DF$Inv.issue[i] <- 1 }
            
          }
        }
        
      }
  } else if ( DF$Publication[i] >= Sys.Date() & !is.na(DF$Publication[i]) ){
    DF$WOH[i] <- NaN
  }
}


#Amazon WOH calculation----------------------------------------------------------

DF$AMZ_inv_temp <- DF$`Retail inv`
DF$WOH_AMZ <- '12+'

for (i in 1:nrow(DF)){
  
  #If the sum of projected sales is lower than available inventory, not published yet, and non null
  if ( !is.na(DF$`Retail inv`[i]) &  DF$Publication[i] < Sys.Date() & 
       !is.na(DF$Publication[i]) & DF$`Retail inv`[i]  <= as.numeric(rowSums(DF[i,10:21])) ){ 
    
    a = 1
    b = 1
    temp_value <- DF$`Retail inv`[i]
    
    for ( j in 10:21){
      
      if ( DF$AMZ_inv_temp[i] - DF[i,j] > 0 ){
        
        DF$AMZ_inv_temp[i] <- DF$AMZ_inv_temp[i] - DF[i,j]
        temp_value <-  DF$AMZ_inv_temp[i]
        b = b + 1
        
      } else if ( DF$AMZ_inv_temp[i] - DF[i,j]  <= 0){
        
        if (a == 1){
          
          DF$WOH_AMZ[i] <- b
          a = a + 1 }
        
        
      }
    }
    
  }  else if ( DF$Publication[i] >= Sys.Date() & !is.na(DF$Publication[i]) ){
    DF$WOH_AMZ[i] <- NaN }

}


DF <- merge(DF, AA_status, by = "isbn", all.x = TRUE)

DF <- merge(DF, Print_status, by = "isbn", all.x = TRUE)


pred_df_holt_damp_beta <- DF

#-----------------------------------------------------------------------------------------------------------------------
#                                 Cleaning and re-ordering
#-----------------------------------------------------------------------------------------------------------------------



#Replacing negative predictions with 0 and rounding all predictions to closest number
for ( i in 1:nrow(pred_df_holt_damp_beta)){
  
  for ( j in 10:21){
    
    
    pred_df_holt_damp_beta[i,j] <- round(pred_df_holt_damp_beta[i,j], digits = 0)
    
    if (pred_df_holt_damp_beta[i,j]<0){
      
      pred_df_holt_damp_beta[i,j] <- 0
      
    }
  }  
}

#Adding inventory info
pred_df_holt_damp_beta$Inventory <- pred_df_holt_damp_beta$`Retail inv` + pred_df_holt_damp_beta$`TBS inv` #+ pred_df_holt_damp_beta$`AMZ Open Orders`


#Computing slope
for (i in 1:nrow(pred_df_holt_damp_beta)){
  
  pred_df_holt_damp_beta$Slope[i] <- round(as.numeric(lm(as.numeric(pred_df_holt_damp_beta[i,c(6:9)]) ~ c(1:4))$coefficients[2]),1)
  
}



pred_df_holt_damp_beta$'12w_inventory_adjusted' <- rowSums(pred_df_holt_damp_beta[,c(10:21)])



pred_df_holt_damp_beta <- arrange(pred_df_holt_damp_beta, desc(pred_df_holt_damp_beta[,9]))

pred_df_holt_damp_beta$Publication <- as.Date(pred_df_holt_damp_beta$Publication)

pred_df_holt_damp_beta$`Reprint Qty`[is.na(pred_df_holt_damp_beta$`Reprint Qty`)] <- NaN

pred_df_holt_damp_beta <- pred_df_holt_damp_beta %>%
  relocate(`Reprint Qty`, .after = `12w_inventory_adjusted`)

#re-ordering columns
pred_df_holt_damp_beta <- pred_df_holt_damp_beta[,c(2,1,3:22,24:27,34,33,35,29,30:32,36,37,38)]

forecasting_statistics_temp <- pred_df_holt_damp_beta[,c(1:22,35,34,31,33,23,24,29,30,25,36)]
pred_df_holt_damp_beta <- pred_df_holt_damp_beta[,c(1:22,35,34,31,33,23,24,29,30,25,27,28,26,36)]

Hitlist_titles <- c("0744036720","1465449817","0744020530","1465453237","1465436030","0744020565","1465474765",
                    "146546848X","1465463305","146544968X","074402997X","1465475850","1465437975","0744035015",
                    "1465436022","1465451439","1465474935","1465497897",
                    "1465478906","1465436073","1465499334","0744020573","074402885X","1615649980","0744057051")


ignore_list <- c("9781465477354", "9781465483676", "9780756673178", "9781465445483", "9780241407752", "9781465478788",
                 "9780241568583", "9780241407721", "9780241462836")




# --------------------------------------------------------------------------------------------------------------
#                               Saving file + formatting                                                        \
# --------------------------------------------------------------------------------------------------------------

# Adds after the second column
pred_df_holt_damp_beta <- add_column(pred_df_holt_damp_beta, new_col = NA, .after = 5)
pred_df_holt_damp_beta <- add_column(pred_df_holt_damp_beta, new_col = NA, .after = 22)
pred_df_holt_damp_beta <- add_column(pred_df_holt_damp_beta, new_col = NA, .after = 28)
pred_df_holt_damp_beta <- add_column(pred_df_holt_damp_beta, new_col = NA, .after = 33)

colnames(pred_df_holt_damp_beta)[3] <- "Title"
colnames(pred_df_holt_damp_beta)[6] <- ""
colnames(pred_df_holt_damp_beta)[23] <- ""
colnames(pred_df_holt_damp_beta)[29] <- ""
colnames(pred_df_holt_damp_beta)[34] <- ""


pred_df_holt_damp_beta$isbn <- as.character(pred_df_holt_damp_beta$isbn)
pred_df_holt_damp_beta$asin <- as.character(pred_df_holt_damp_beta$asin)



#Highlighting columns with inventory issues
wb <- createWorkbook()
addWorksheet(wb, sheetName="US")
writeData(wb, sheet="US", x=pred_df_holt_damp_beta)



#adding filters
addFilter(wb, "US", rows = 1, cols = 1:ncol(pred_df_holt_damp_beta))

#auto width for columns
width_vec <- suppressWarnings(apply(pred_df_holt_damp_beta, 2, function(x) max(nchar(as.character(x)) + 1, na.rm = TRUE)))
width_vec_header <- nchar(colnames(pred_df_holt_damp_beta))  + 3
max_vec_header <- pmax(width_vec, width_vec_header)
setColWidths(wb, "US", cols = 1:ncol(pred_df_holt_damp_beta), widths = max_vec_header )
setColWidths(wb, "US",  cols = 1, widths = 13)
setColWidths(wb, "US",  cols = 2, widths = 15)
setColWidths(wb, "US",  cols = 3, widths = 52)
setColWidths(wb, "US",  cols = 27, widths = 7)
setColWidths(wb, "US",  cols = 6, widths = 10)
setColWidths(wb, "US",  cols = 23, widths = 10)
setColWidths(wb, "US",  cols = 29, widths = 10)
setColWidths(wb, "US",  cols = 34, widths = 10)




# Highlighting rows
yellow_style <- createStyle(fgFill="#FFFF00")
x <- which( pred_df_holt_damp_beta$Inv.issue ==1 
            &  pred_df_holt_damp_beta$Print_status %notin% c( "No Reprints Sched.", "Out of Stock Indef.", "Out of Print", "Remainder" ) 
            & (pred_df_holt_damp_beta$Reprint.Date >= Sys.Date() + 100 | is.na(pred_df_holt_damp_beta$Reprint.Date)  ) 
            & pred_df_holt_damp_beta$isbn %notin% ignore_list )

addStyle(wb, sheet="US", style=yellow_style, rows=x+1, cols=c(1:ncol(pred_df_holt_damp_beta)), 
         gridExpand=TRUE, stack = TRUE) # +1 for header line


#Centering cells
centerStyle <- createStyle(halign = "center")
addStyle(wb, "US", style=centerStyle, rows = 2:nrow(pred_df_holt_damp_beta), cols = 7:ncol(pred_df_holt_damp_beta), 
         gridExpand = T, stack = TRUE)

# Adding borders
invisible(OutsideBorders(
  wb,
  sheet_ = "US",
  rows_ = 1:nrow(pred_df_holt_damp_beta),
  cols_ = 1:5
))

invisible(OutsideBorders(
  wb,
  sheet_ = "US",
  rows_ = 1:nrow(pred_df_holt_damp_beta),
  cols_ = 7:10
))

invisible(OutsideBorders(
  wb,
  sheet_ = "US",
  rows_ = 1:nrow(pred_df_holt_damp_beta),
  cols_ = 11:22
))

invisible(OutsideBorders(
  wb,
  sheet_ = "US",
  rows_ = 1:nrow(pred_df_holt_damp_beta),
  cols_ = 24:28
))

invisible(OutsideBorders(
  wb,
  sheet_ = "US",
  rows_ = 1:nrow(pred_df_holt_damp_beta),
  cols_ = 30:33
))

invisible(OutsideBorders(
  wb,
  sheet_ = "US",
  rows_ = 1:nrow(pred_df_holt_damp_beta),
  cols_ = 35:39
))

freezePane(
  wb,
  sheet = "US",
  firstActiveRow = 2,
  firstActiveCol = 6
)

pred_date <- Sys.Date() + 6 - match(weekdays(Sys.Date()), all_days)

saveWorkbook(wb, paste0("Bookscan Forecast us - ",pred_date,".xlsx"), overwrite = T) 


cat("\nForecast end\n")
