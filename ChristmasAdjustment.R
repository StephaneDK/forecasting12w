library.path <- .libPaths("C:/Users/steph/Documents/R/win-library/4.0")
#library.path <- .libPaths("C:\\Users\\snichanian\\Documents\\R\\R-4.0.3\\library")

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


#-----------------------------------------------------------------------------------------------------------------------
#                                 Christmas period adjustment
#-----------------------------------------------------------------------------------------------------------------------

ChristmasAdjustment <- function(pred_df_holt_damp_beta, Q1){

#Subsetting to Q4-2020 sales
xmas_sales <- Q1
xmas_sales <- xmas_sales[,c(1:5,grep(as.character(as.Date(colnames(pred_df_holt_damp_beta[6])) - 364), colnames(xmas_sales) ):
                              grep(as.character(as.Date(colnames(pred_df_holt_damp_beta[21])) - 364), colnames(xmas_sales)) )]
xmas_sales <- xmas_sales[xmas_sales$Publication <= "2021-10-01",]

#Deleting seasonal, highly influential titles
xmas_sales <- xmas_sales[!(xmas_sales$isbn %in% c("9781465462350", "9781465452764", "9781465483614") ),]

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


#Seasonal adjustment loop (only between Nov 08 2021 and Jan 15 2022)
for (i in 1:nrow(pred_df_holt_damp_beta)){
  
  if (pred_df_holt_damp_beta$isbn[i] %notin% c("9781465462350", "9781465452764", "9781465483614") ){
    
    for (j in 10:21){
      
      if (as.Date(colnames(pred_df_holt_damp_beta)[j]) >= "2022-11-08" &
          as.Date(colnames(pred_df_holt_damp_beta)[j]) <= "2023-01-15" )
        
        pred_df_holt_damp_beta[i,j] <- ceiling(pred_df_holt_damp_beta[i,j-1]*
                                                 xmas_change_per[ xmas_change_per$Group.1 == pred_df_holt_damp_beta$Division[i],
                                                                  as.character(as.Date(colnames(pred_df_holt_damp_beta)[j]) - 364) ] )
      
    }
  }
}

pred_original <- pred_df_holt_damp_beta

return(pred_original)

}

