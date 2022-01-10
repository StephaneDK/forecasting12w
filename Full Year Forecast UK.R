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
suppressMessages(library(magrittr, lib.loc = library.path))

options(scipen=999, digits = 3, error=function() { traceback(2); if(!interactive()) quit("no", status = 1, runLast = FALSE) } )
all_days <- c("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")

`%notin%` <- Negate(`%in%`)
current_quarter <- "Q3"

#Setting the directory where all files will be used from for this project
setwd("C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv")

#Importing Sales Data
Sales_UK <- read.csv("Sales uk.csv", header = T, stringsAsFactors = FALSE)
Sales_UK <- Sales_UK %>% 
            mutate(date = as.Date(Sales_UK$date)) %>% 
            subset(title != "")


#Importing Reprint and AA Data
Reprint <- read.csv("UK Inventory.csv", header = T, stringsAsFactors = FALSE)
Reprint <- Reprint %>%
           select(ASIN, 
                  Available.Inventory, 
                  Stock.at.TBS, 
                  Reprint.Dates, 
                  Open.Purchase.Order.Quantity, 
                  B3.Status, 
                  contains("Focus.List.Current.Status"),
                  Reprint.Quantity 
                  ) %>%
           set_colnames(c("asin", 
                          "Amz inv", 
                          "TBS inv",  
                          "Reprint.Date", 
                          "AMZ Open Orders", 
                          "Print Status",
                          "AA Status",
                          "Reprint Qty"))  %>%
           mutate(Reprint.Date = as.Date(Reprint$Reprint.Date, format = "%d/%m/%Y"),
                  `Amz inv` =  suppressWarnings(as.numeric(gsub(",","",`Amz inv`))),
                  `TBS inv` =  suppressWarnings(as.numeric(gsub(",","",`TBS inv`))),
                  `AMZ Open Orders` =  suppressWarnings(as.numeric(gsub(",","",`AMZ Open Orders`))),
                  `Reprint Qty` = replace(`Reprint Qty`, is.na(`Reprint Qty`), NaN)
                  )


#Matching reprint delivery date with closest Saturday
Reprint$Reprint.Date <- Reprint$Reprint.Date + 6 - match(weekdays(Reprint$Reprint.Date), all_days)

cat("hello")

for (i in 1:length(Sales_UK$asin)){
  if (nchar(Sales_UK$asin[i]) == 9){
    Sales_UK$asin[i] <- paste0("0",Sales_UK$asin[i])  
  }
  
}



for (i in 1:length(Reprint$asin)){
  if (nchar(Reprint$asin[i]) == 9){
    Reprint$asin[i] <- paste0("0",Reprint$asin[i])  
  }
  
}






#Creating the Data Frame with all sales data ------------------------------------------------------------------------------------
DF <- cbind.data.frame(Sales_UK$date, Sales_UK$asin, Sales_UK$isbn, Sales_UK$title,Sales_UK$division ,Sales_UK$pub_date,Sales_UK$units ,stringsAsFactors = FALSE)
colnames(DF) <- c("date","asin","isbn","title","Division","Publication","units")

#Rearrange data with dates as columns
Q1 <- dcast(DF, asin + isbn + title + Division + Publication ~ date, value.var="units", fun.aggregate = sum)

#Re-organising by latest saturday
all_days <- c("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
Sat_date <- Sys.Date() -1 - match(weekdays(Sys.Date()), all_days)
Q1 <- arrange(Q1, desc( Q1[,grep(Sat_date,colnames(Q1))]) )



#Detection of seasonal titles --------------------------------------------------------------------------------------------------

Q_iso <- readRDS("Q_iso_UK.csv")

#Q_iso <- rbind.data.frame(Q_iso, c("0241459001","9780241459003","Baby's First Easter","Q2")  )


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
#                                 Seasonal titles adjustment
#-----------------------------------------------------------------------------------------------------------------------


#Create 2020 only data frame
prev_year <- Q1[ Q1$asin %in% Q_iso$asin, ]
prev_year <- prev_year[,c(1,3,grep("2020-01-04", colnames(Q1)): grep("2021-04-10", colnames(Q1))) ]
prev_year[prev_year <= 0] <- 1


#Create empty data frame to store seasonal percentage changes
Seas_adjQ <- data.frame(matrix(ncol = 69 , nrow = nrow(prev_year)))
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
Seas_adjQ[Seas_adjQ$asin == "0241377978",3:16]  <- c( 1, 1.6, 1, 1.5, 1.2, 1.2, 1.5, 0.27, 0.44, 0.75, 0.66, 1, 1, 1 )

#manual adjustment on existing seasonal titles
Seas_adjQ[Seas_adjQ$asin == "0241283477",3:16]  <- c(1, 0.76, 0.83, 1.2, 1.65, 1.18, 1.60, 0.18,  0.99, 0.96, 1.02, 1.38, 1.02, 1.47) 

#Rounding to 3 decimal places
Seas_adjQ[,3:ncol(Seas_adjQ)] <- round(Seas_adjQ[,3:ncol(Seas_adjQ)],3)



#Keep only current seasonal adjustment data
Seas_adjQ <- Seas_adjQ[ Seas_adjQ$asin %in% subset(Q_iso, Q_iso$Season == current_quarter)$asin ,]

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



#Category Adjustment -----------------------------------------------------------------------------------------------------------------

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
#                                 Christmas period adjustment
#-----------------------------------------------------------------------------------------------------------------------


#Subsetting to Q4-2020 sales
xmas_sales <- Q1 
xmas_sales <- xmas_sales[,c(1:5,grep(as.character(as.Date(colnames(pred_df_holt_damp_beta[6])) - 364), colnames(xmas_sales) ):
                              grep(as.character(as.Date(colnames(pred_df_holt_damp_beta[21])) - 364), colnames(xmas_sales)) )]
xmas_sales <- xmas_sales[xmas_sales$Publication <= "2020-10-01",]


#Division aggregates
xmas_change <- aggregate(xmas_sales[,6:21], list(xmas_sales$Division), mean)

#Create empty data frame to store seasonal percentage changes
xmas_change_per <- data.frame(matrix(ncol = ncol(xmas_change) , nrow = nrow(xmas_change)) ) #Same Structure
colnames(xmas_change_per) <- colnames(xmas_change) #Same column names
xmas_change_per$Group.1 <- xmas_change$Group.1 #Same Division names


#Compute seasonal percentage changes using last year data
for (i in 1:nrow(xmas_change_per)){
  
  for ( j in 1:ncol(xmas_change_per)){
    
    if (j >= 2 & j <=6 ){
      xmas_change_per[i,j] <- 1
      
    } else if (j > 6) {
      xmas_change_per[i,j] <- xmas_change[i,j] / xmas_change[i,j-1]
      
    }
    
  }
  
}

xmas_change_per[is.na(xmas_change_per)] <- 1


#Seasonal adjustment loop (only between Nov 08 2021 and Jan 15 2022)
for (i in 1:nrow(pred_df_holt_damp_beta)){
  
  
  for (j in 10:21){
    
    if (as.Date(colnames(pred_df_holt_damp_beta)[j]) >= "2021-11-13" &
        as.Date(colnames(pred_df_holt_damp_beta)[j]) <= "2022-01-15" )
      
      pred_df_holt_damp_beta[i,j] <- ceiling(pred_df_holt_damp_beta[i,j-1]*
                                               xmas_change_per[ xmas_change_per$Group.1 == pred_df_holt_damp_beta$Division[i], 
                                                                as.character(as.Date(colnames(pred_df_holt_damp_beta)[j]) - 364) ] )
    
  }
  
}



#Title level adjustment for top 20 titles --------------------------------------

pred_df_holt_damp_beta <- arrange(pred_df_holt_damp_beta, desc(pred_df_holt_damp_beta[,9]) )

#Create 2020 only data frame
prev_year <- Q1[ Q1$asin %in% pred_df_holt_damp_beta[1:20,1], ]
prev_year <- prev_year[,c(1,3,grep("2020-01-04", colnames(Q1)): grep("2021-04-10", colnames(Q1))) ]
prev_year[prev_year <= 0] <- 1


#Create empty data frame to store seasonal percentage changes
Seas_adjQ <- data.frame(matrix(ncol = 69 , nrow = nrow(prev_year)))
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


#Rounding to 3 decimal places
Seas_adjQ[,3:ncol(Seas_adjQ)] <- round(Seas_adjQ[,3:ncol(Seas_adjQ)],3)

Seas_adjQ <- arrange(Seas_adjQ, desc(Seas_adjQ$asin))
temp <- arrange(pred_df_holt_damp_beta[1:20,], desc(pred_df_holt_damp_beta[1:20,1]) )


#Seasonal adjustment loop
for (i in 1:nrow(temp)){
  
  if (temp$asin[i] %in% Seas_adjQ$asin && temp$Publication[i] <= '2020-09-10' && 
      temp$asin[i] %notin% c("1409366553","1405393505","0241455154","0241352487","0241287936")){
    
    for (j in 10:21){
      
      temp[i,j] <- temp[i,j-1]*Seas_adjQ[ Seas_adjQ$asin == temp$asin[i], 
                                          as.character(as.Date(colnames(temp)[j]) - 364) ]
      
    }
  }
}

pred_original <- pred_df_holt_damp_beta

#assigning new values to pred DF
for (i in 1:nrow(temp)){
  
  pred_df_holt_damp_beta[which(pred_df_holt_damp_beta$asin == temp$asin[i] ),] <- temp[i,] 
  
}



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
y <- which( colnames(pred_df_holt_damp_beta)=="Inv_issue" )
x <- which( pred_df_holt_damp_beta$Inv_issue ==1 )
addStyle(wb, sheet="UK", style=yellow_style, rows=x+1, cols=c(1:ncol(pred_df_holt_damp_beta)), 
         gridExpand=TRUE, stack = TRUE) # +1 for header line

# Highlighting rows in light green for Hitlist
green_style <- createStyle(fgFill="#90EE90")
x <- which( pred_df_holt_damp_beta$asin %in% Hitlist_titles )
addStyle(wb, sheet="UK", style=green_style, rows=x+1, cols=c(1:ncol(pred_df_holt_damp_beta)), 
         gridExpand=TRUE, stack = TRUE) # +1 for header line

# Highlighting rows in orange for Hitlist issue
orange_style <- createStyle(fgFill="#FF4500")
x <- which( pred_df_holt_damp_beta$asin %in% Hitlist_titles &  pred_df_holt_damp_beta$Inv_issue ==1)
addStyle(wb, sheet="UK", style=orange_style, rows=x+1, cols=c(1:ncol(pred_df_holt_damp_beta)), 
         gridExpand=TRUE, stack = TRUE) # +1 for header line


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



