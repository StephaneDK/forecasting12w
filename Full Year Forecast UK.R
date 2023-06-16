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
Sales_UK <- read.csv("Sales uk.csv", header = T, stringsAsFactors = FALSE)
Sales_UK <- Sales_UK %>% 
  rename_all(tolower)  %>% 
  subset(title != "" ) %>% 
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

Q1 <- Sales_UK

#Import AA Data
AA_status <-  read.csv("AA Status uk.csv", header = T, stringsAsFactors = FALSE)
colnames(AA_status) <- c("ISBN", "AA_status")

#Import Print Status Data
Print_status <-  read.csv("Print Status uk.csv", header = T, stringsAsFactors = FALSE)
colnames(Print_status) <- c(  "ISBN", "Print_status")


#Importing reprint dates & quantity
Reprint2 <- read.csv("Reprint.csv", header = T, stringsAsFactors = FALSE) 
#sort(colnames(Reprint2))
Reprint2 <- Reprint2 %>% 
  subset(Print.Instruction.Impression %in% c("DK UK" ) ) %>% 
  mutate(Print.Instruction.Requested.Delivery.Date = as.Date(Print.Instruction.Requested.Delivery.Date, format = "%d/%m/%Y")) %>% 
  select( ISBN, Print.Instruction.Requested.Delivery.Date,  Print.Instruction.Quantity) %>% 
  group_by(ISBN) %>%
  filter(Print.Instruction.Requested.Delivery.Date == min(Print.Instruction.Requested.Delivery.Date) )  %>%
  top_n(1, Print.Instruction.Quantity) %>%
  set_colnames(c( "ISBN", "Reprint.Date", "Reprint Qty" ))


#Importing Inventory
Inventory <- read.csv("Inventory uk.csv", header = T, stringsAsFactors = FALSE) 
#sort(colnames(Reprint2))

Inventory <- Inventory %>% 
  select( ASIN, MATERIAL, AMZ_INV, TBS_INV, AMZ_OPEN_ORDERS) %>% 
  set_colnames(c("asin" ,"ISBN", "Amz inv", "TBS inv", "AMZ Open Orders" )) %>%
  merge( Reprint2, by = "ISBN", all.x = T) %>%
  merge( AA_status, by = "ISBN", all.x = T) %>%
  merge( Print_status, by = "ISBN", all.x = T) %>%
  select(asin, ISBN, `Amz inv`, `TBS inv`, Reprint.Date, 
         `AMZ Open Orders`, Print_status, AA_status,
         `Reprint Qty`) %>%
  dplyr::rename("Print Status" = "Print_status",
                "AA Status" = "AA_status") %>%
  mutate(Reprint.Date = Reprint.Date + 6 - match(weekdays(Reprint.Date), all_days),
         `TBS inv` = ifelse(is.na(`TBS inv`), 0, `TBS inv`),
         `Amz inv` =  ifelse(is.na(`Amz inv`), 0, `Amz inv`),
         `Reprint Qty` = replace(`Reprint Qty`, is.na(`Reprint Qty`), NaN),
         ISBN = NULL)
  
Reprint <- Inventory


#Import seasonal titles --------------------------------------------------------------------------------------------------

Q_iso <- readRDS("Q_iso_UK.csv")


#Adding seasonal titles manually (if missing)
Q_iso <- rbind.data.frame(Q_iso, c("0241287790","9780241287798","Baby Touch and Feel: Halloween","Q4")  )
Q_iso <- rbind.data.frame(Q_iso, c("0241484340","9780241484340","The Happy Pumpkin","Q4")  )



#Creating computational statistics ---------------------------------------------------------------------------------------------------

#Time Vectors
time_vec <- c(1:4)
time_vec_future <- c(5:16)

#Training and forecasting sizes
training_size <- ncol(Q1)-182+1
forecast_size <- 14 - (ncol(Q1)-182+1) 




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

if ( ncol(b) > 2 ){
  pred_df_holt_damp_beta <- cbind.data.frame(Q1[,c(1:5,(ncol(Q1) - 3):ncol(Q1))],a, b[,3:ncol(b)])
  
} else {
  pred_df_holt_damp_beta <- cbind.data.frame(Q1[,c(1:5,182:ncol(Q1))],a)
  
}


#Renaming columns

for (i in 1:ncol(pred_df_holt_damp_beta)){
 
  if (i > 9){

    colnames(pred_df_holt_damp_beta)[i] <-  as.character( as.Date(colnames(pred_df_holt_damp_beta)[i-1]) + 7 )
  }
}


#Adjusting for predictions with big negative slopes ----------------------------

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
Seas_adjQ <- data.frame(matrix(ncol =  ncol(prev_year) , nrow = nrow(prev_year)))
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

#manual adjustment on existing seasonal titles
#Seas_adjQ[Seas_adjQ$asin == "0241377978",3:16]  <- c( 1, 1.6, 1, 1.5, 1.2, 1.2, 1.5, 0.27, 0.44, 0.75, 0.66, 1, 1, 1 )
#Seas_adjQ[Seas_adjQ$asin == "0241283477",3:16]  <- c(1, 0.76, 0.83, 1.2, 1.65, 1.18, 1.60, 0.18,  0.99, 0.96, 1.02, 1.38, 1.02, 1.47) 

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
                                                                              as.character(as.Date(colnames(pred_df_holt_damp_beta)[j]) - 364) ]
      
    }
  }
}


#-----------------------------------------------------------------------------------------------------------------------
#                                 Category Adjustment 
#-----------------------------------------------------------------------------------------------------------------------

# #Last 4 weeks of children book are lowered by 20%
# for ( i in 1:nrow(pred_df_holt_damp_beta)){
#   if (pred_df_holt_damp_beta$Division[i] == "Children 0-9" | pred_df_holt_damp_beta$Division[i] == "Knowledge Children"){
#     pred_df_holt_damp_beta[i,16:19] <-  pred_df_holt_damp_beta[i,16:19]*0.8 
#   }
# }


#Last 7 weeks of travel books are increased by 20%
for ( i in 1:nrow(pred_df_holt_damp_beta)){
  if (pred_df_holt_damp_beta$Division[i] == "Travel"){
    pred_df_holt_damp_beta[i,10:21] <-  pred_df_holt_damp_beta[i,10:21]*1.2 
  }
}

#Manual adjustment for books with PR campaigns or sudden sales spikes (not used now)




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


#-----------------------------------------------------------------------------------------------------------------------
#                                 Inventory Adjustment
#-----------------------------------------------------------------------------------------------------------------------

#Calculating totals
pred_df_holt_damp_beta$Total_12w <- rowSums(pred_df_holt_damp_beta[,c(10:21) ])


#Merge reprint & Sales Data Frames
DF <- merge(pred_df_holt_damp_beta, Reprint, by = "asin", all.x = TRUE)

#Replace negative values of inventory in TBS by 0 (not meaningful)
for ( i in 1:nrow(DF)){
  
  if (DF$`TBS inv`[i] < 0 & !is.na(DF$`TBS inv`[i])){
    DF$`TBS inv`[i] <- 0
  }
  
  if (DF$`Amz inv`[i] < 0 & !is.na(DF$`Amz inv`[i])){
    DF$`Amz inv`[i] <- 0
  }
  
}


#One variable for amazon and tbs inventory
DF$inventory <- DF$`Amz inv` + DF$`TBS inv` # + DF$`AMZ Open Orders` 
DF$Inv_issue <- 0
DF$WOH <- "12+"

DF$Publication <- as.Date(DF$Publication)


#DK Inventory stock re-adjustment
for (i in 1:nrow(DF)){
  
  #If the sum of projected sales is lower than available inventory, not published yet, and non null
  if ( !is.na(DF$inventory[i]) & DF$Publication[i] < Sys.Date() & 
       !is.na(DF$Publication[i]) & DF$inventory[i]  < as.numeric(rowSums(DF[i,10:21])) ){ 
      
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
              DF$Inv_issue[i] <- 1 }
            
          }
        }
        
      }
      
  } else if ( DF$Publication[i] >= Sys.Date() & !is.na(DF$Publication[i]) ){
    DF$WOH[i] <- NaN
  }
}


DF$AMZ_inv_temp <- DF$`Amz inv`
DF$WOH_AMZ <- "12+"


#AMZ WOH calculation
for (i in 1:nrow(DF)){
  
  #If the sum of projected sales is lower than available inventory, not published yet, and non null
  if ( !is.na(DF$`Amz inv`[i]) &  DF$Publication[i] < Sys.Date() & 
       !is.na(DF$Publication[i]) & DF$`Amz inv`[i]  <= as.numeric(rowSums(DF[i,10:21])) ){ 
    
     
      
      a = 1
      b = 1
      temp_value <- DF$`Amz inv`[i]
      
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
        
  }  else if ( DF$Publication[i] >= Sys.Date() & !is.na(DF$Publication[i])){
    DF$WOH_AMZ[i] <- NaN }
      
}


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
pred_df_holt_damp_beta$inventory <- pred_df_holt_damp_beta$`Amz inv` + pred_df_holt_damp_beta$`TBS inv` #+ pred_df_holt_damp_beta$`AMZ Open Orders`


#Computing slope
for (i in 1:nrow(pred_df_holt_damp_beta)){
  
  pred_df_holt_damp_beta$Slope[i] <- round(as.numeric(lm(as.numeric(pred_df_holt_damp_beta[i,c(6:9)]) ~ c(1:4))$coefficients[2]),1)
  
}



pred_df_holt_damp_beta$'12w_inventory_adjusted' <- rowSums(pred_df_holt_damp_beta[,c(10:21)])


pred_df_holt_damp_beta <- arrange(pred_df_holt_damp_beta, desc(pred_df_holt_damp_beta[,9]))

pred_df_holt_damp_beta$Publication <- as.Date(pred_df_holt_damp_beta$Publication)

pred_df_holt_damp_beta <- pred_df_holt_damp_beta %>%
  relocate(`Reprint Qty`, .after = `12w_inventory_adjusted`)

forecasting_statistics_temp <- pred_df_holt_damp_beta[,c(1:22,35,34,31,33,23,24,29,30,25,36)]

pred_df_holt_damp_beta <- pred_df_holt_damp_beta[,c(1:22,35,34,31,33,23,24,29,30,25,27,28,26,36)]

Hitlist_titles <- c("0241315611","0241343267","0241424305","0241357551","0241224934","0241229782","0241412471",
                    "024135871X","0241302323","0241426162","024133439X","0241440610","0241286123","0241446619",
                    "0241515106","0241446341","0241412706")

ignore_list <- c("9780241228371", "9781405328319", "9780241407721", "9780241511107", "9780241509746", "9780241559482",
                 "9780241509678", "9780241533321", "9780241520475", "9781409347965", "9780241327388", "9780241462621",
                 "9781405351782", "9780241520475", "9780241509654", "9780241509586", "9781405336598", "9780241510643",
                 "9781405375818", "9780241462836", "9780751321531", "9780241361979", "9780241368848", "9780241509616",
                 "9780241533307", "9780241510599", "9780241544310", "9780241568897", "9780241598436", "9780241302323",
                 "9780241500866", "9780241250310", "9781405328609")


# --------------------------------------------------------------------------------------------------------------
#                               Saving file + formatting                                                        \
# --------------------------------------------------------------------------------------------------------------

# Adds after the second column
pred_df_holt_damp_beta <- add_column(pred_df_holt_damp_beta, new_col = NA, .after = 5)
pred_df_holt_damp_beta <- add_column(pred_df_holt_damp_beta, new_col = NA, .after = 22)
pred_df_holt_damp_beta <- add_column(pred_df_holt_damp_beta, new_col = NA, .after = 28)
pred_df_holt_damp_beta <- add_column(pred_df_holt_damp_beta, new_col = NA, .after = 33)

colnames(pred_df_holt_damp_beta)[6] <- ""
colnames(pred_df_holt_damp_beta)[23] <- ""
colnames(pred_df_holt_damp_beta)[29] <- ""
colnames(pred_df_holt_damp_beta)[34] <- ""


pred_df_holt_damp_beta$isbn <- as.character(pred_df_holt_damp_beta$isbn)
pred_df_holt_damp_beta$asin <- as.character(pred_df_holt_damp_beta$asin)


#Highlighting columns with inventory issues
wb <- createWorkbook()
addWorksheet(wb, sheetName="UK")
writeData(wb, sheet="UK", x=pred_df_holt_damp_beta)


#adding filters
addFilter(wb, "UK", rows = 1, cols = 1:ncol(pred_df_holt_damp_beta))

#auto width for columns
width_vec <- suppressWarnings(apply(pred_df_holt_damp_beta, 2, function(x) max(nchar(as.character(x)) + 1, na.rm = TRUE)))
width_vec_header <- nchar(colnames(pred_df_holt_damp_beta))  + 3
max_vec_header <- pmax(width_vec, width_vec_header)
setColWidths(wb, "UK", cols = 1:ncol(pred_df_holt_damp_beta), widths = max_vec_header )
setColWidths(wb, "UK",  cols = 1, widths = 13)
setColWidths(wb, "UK",  cols = 2, widths = 15)
setColWidths(wb, "UK",  cols = 3, widths = 52)
setColWidths(wb, "UK",  cols = 27, widths = 7)
setColWidths(wb, "UK",  cols = 6, widths = 10)
setColWidths(wb, "UK",  cols = 23, widths = 10)
setColWidths(wb, "UK",  cols = 23, widths = 9)
setColWidths(wb, "UK",  cols = 29, widths = 10)
setColWidths(wb, "UK",  cols = 34, widths = 10)



# Highlighting rows in yellow for inventory issue
yellow_style <- createStyle(fgFill="#FFFF00")
orange_style <- createStyle(fgFill="#FFA500")


#x <- which( pred_df_holt_damp_beta$WOH_AMZ %in% c(1,2)  & pred_df_holt_damp_beta$`Print Status`!= "Out of Print")
#addStyle(wb, sheet="UK", style=orange_style, rows=x+1, cols=c(1:ncol(pred_df_holt_damp_beta)), 
         #gridExpand=TRUE, stack = TRUE) # +1 for header line


x <- which( pred_df_holt_damp_beta$Inv_issue == 1  
            & pred_df_holt_damp_beta$`Print Status`!= "Out of Print"
            & (pred_df_holt_damp_beta$Reprint.Date >= Sys.Date() + 100 | is.na(pred_df_holt_damp_beta$Reprint.Date) )
            & pred_df_holt_damp_beta$isbn %notin% ignore_list)

addStyle(wb, sheet="UK", style=yellow_style, rows=x+1, cols=c(1:ncol(pred_df_holt_damp_beta)), 
         gridExpand=TRUE, stack = TRUE) # +1 for header line


# # Highlighting rows in light green for Hitlist
# green_style <- createStyle(fgFill="#90EE90")
# x <- which( pred_df_holt_damp_beta$asin %in% Hitlist_titles )
# addStyle(wb, sheet="UK", style=green_style, rows=x+1, cols=c(1:ncol(pred_df_holt_damp_beta)), 
#          gridExpand=TRUE, stack = TRUE) # +1 for header line
# 
# # Highlighting rows in orange for Hitlist issue
# orange_style <- createStyle(fgFill="#FF4500")
# x <- which( pred_df_holt_damp_beta$asin %in% Hitlist_titles &  pred_df_holt_damp_beta$Inv_issue ==1)
# addStyle(wb, sheet="UK", style=orange_style, rows=x+1, cols=c(1:ncol(pred_df_holt_damp_beta)), 
#          gridExpand=TRUE, stack = TRUE) # +1 for header line


#Centering cells
centerStyle <- createStyle(halign = "center")
addStyle(wb, "UK", style=centerStyle, rows = 2:nrow(pred_df_holt_damp_beta), cols = 7:ncol(pred_df_holt_damp_beta), 
         gridExpand = T, stack = TRUE)

# Adding borders
invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(pred_df_holt_damp_beta),
  cols_ = 1:5
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(pred_df_holt_damp_beta),
  cols_ = 7:10
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(pred_df_holt_damp_beta),
  cols_ = 11:22
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(pred_df_holt_damp_beta),
  cols_ = 24:28
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(pred_df_holt_damp_beta),
  cols_ = 30:33
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(pred_df_holt_damp_beta),
  cols_ = 35:39
))

freezePane(
  wb,
  sheet = "UK",
  firstActiveRow = 2,
  firstActiveCol = 6
)


pred_date <- Sys.Date() + 6 - match(weekdays(Sys.Date()), all_days)

saveWorkbook(wb, paste0("Forecast uk - ",pred_date,".xlsx"), overwrite = T) 

#File Formatting for Slowing Trending ----------------------------------------------------------------------------------------------------
temp <- pred_df_holt_damp_beta[,1:33]
colnames(temp)[6] <- ""
colnames(temp)[23] <- ""
colnames(temp)[29] <- ""

write.csv(temp, paste0("Forecast uk - ",pred_date,".csv"), row.names = F)

cat("\n\nUK Forecasts saved succesfully\n")


#File Formatting for DB write ----------------------------------------------------------------------------------------------
pred_raw <- pred_df_holt_damp_beta[,c(1,11:22)]


pred_formatted <- melt(pred_raw, value.name = "pred_units", id.vars = "asin")
pred_formatted$pred_at <- as.character(Sys.Date())
pred_formatted$region <- "uk"
pred_formatted$model <- "Holt"


pred_formatted <- pred_formatted[,c(5,6,1,4,2,3)]

colnames(pred_formatted) <- c("REGION","MODEL","ASIN","PRED_AT","PRED_FOR","PRED_UNITS")


write.csv(pred_formatted, "Q1 Forecast uk - Formatted.csv",row.names = F, quote = FALSE)

cat("UK Forecasts formatted succesfully\n")


#File Formatting for DB write ----------------------------------------------------------------------------------------------
write_db <- forecasting_statistics_temp



write_db$pred_at <- as.character(Sys.Date())
write_db$region <- "uk"
write_db$model <- "Holt"

write_db <- write_db[,c(34,35,1,33,27:29,22,23,25,26,30,31,32)]

colnames(write_db) <- c("region","model","asin","pred_at","Amz_inv","DK_inv", "Total_inventory", 
                        "Forecast_12w", "Adjusted_forecast_12w", "Weeks_on_hand", "Weeks_on_hand_AMZ",
                        "Inv_issue","reprint_date","reprint_quantity")


for (i in 5:9){
  write_db[,i] <- as.integer(write_db[,i])
  write_db[is.na(write_db[,i]),i] <-0
  
}

write_db[is.na(write_db$reprint_date),13] <- "2020-01-01"
write_db$reprint_quantity <- as.character(write_db$reprint_quantity)
write_db$reprint_quantity[is.na(write_db$reprint_quantity)] <- "NaN"

write.csv(write_db, "Q1 Forecast uk.csv", row.names = FALSE, quote = FALSE)


#File Formatting for Blacklist titles ----------------------------------------------------------------------------------------------
pred_df_holt_damp_beta$`Reprint Qty` <- NULL
saveRDS(pred_df_holt_damp_beta, "pred_df_holt_damp_beta_UK.rds")

cat("\nForecast end\n")

