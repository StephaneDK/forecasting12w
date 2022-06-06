library.path <- .libPaths("C:/Users/steph/Documents/R/win-library/4.0")

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
#                                 Title level adjustment for top 55 titles
#-----------------------------------------------------------------------------------------------------------------------



TopTitlesAdjustment <- function(pred_df_holt_damp_beta, Q1){


pred_df_holt_damp_beta <- arrange(pred_df_holt_damp_beta, desc(pred_df_holt_damp_beta[,9]) )

#Create 2020 only data frame
prev_year <- Q1[ Q1$asin %in% pred_df_holt_damp_beta[1:55,1], ]
prev_year <- prev_year[,c(1,3,grep("2021-01-02", colnames(Q1)): grep("2022-01-01", colnames(Q1))) ]
prev_year[prev_year <= 0] <- 1

#Create empty data frame to store seasonal percentage changes
Seas_adjQ <- data.frame(matrix(ncol = 55 , nrow = nrow(prev_year)))
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
   temp <- arrange(pred_df_holt_damp_beta[1:55,], desc(pred_df_holt_damp_beta[1:55,1]) )
  
  
   #Seasonal adjustment loop
   for (i in 1:nrow(temp)){
  
     if (temp$asin[i] %in% Seas_adjQ$asin && temp$Publication[i] <= '2020-09-10' &&
         temp$asin[i] %notin% c("146540855X","1465488820","1465448667","1465468137",
                                "1465479368","1465499334","1465451439","1465476679",
                                "1465478906", "1465447628","1465486828")){
  
       for (j in 10:21){
  
         temp[i,j] <- temp[i,j-1]*Seas_adjQ[ Seas_adjQ$asin == temp$asin[i],
                                             as.character(as.Date(colnames(temp)[j]) - 364) ]
  
       }
     }
   }
  
  pred_original <- pred_df_holt_damp_beta

  #assigning new values to pred DF
   for (i in 1:nrow(temp)){
  
     pred_original[which(pred_original$asin == temp$asin[i] ),] <- temp[i,]
  
   }
  
  

return(pred_original)

}

