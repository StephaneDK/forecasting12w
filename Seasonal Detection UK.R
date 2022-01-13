library.path <- .libPaths("C:/Users/Steph/Documents/R/win-library/4.0")

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


`%notin%` <- Negate(`%in%`)
current_quarter <- "Q3"

#Setting the directory where all files will be used from for this project
setwd("C:\\Users\\Steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv")

#Importing Data
Sales_UK <- read.csv("Sales uk.csv", header = T, stringsAsFactors = FALSE)
Sales_UK$date <- as.Date(Sales_UK$date)
Sales_UK <- Sales_UK[Sales_UK$title != "",]



#Adding 2 weeks for reprint delivery time and matching with closest saturday
all_days <- c("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")


for (i in 1:length(Sales_UK$asin)){
  if (nchar(Sales_UK$asin[i]) == 9){
    Sales_UK$asin[i] <- paste0("0",Sales_UK$asin[i])  
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
Q1S <- Q1
Q1S[is.na(Q1S)] <- 0
Q1S[Q1S<0] <- 0


#R chart method
seas_vec_Q1 <- c()
seas_vec_Q2 <- c()
seas_vec_Q3 <- c()
seas_vec_Q4 <- c()


for (i in 1:nrow(Q1S)){
  
  #2019
  a1 <- as.numeric(Q1S[i,grep("2019-01-05", colnames(Q1S)):grep("2019-03-30", colnames(Q1S))])
  a2 <- as.numeric(Q1S[i,grep("2019-04-06", colnames(Q1S)):grep("2019-06-29", colnames(Q1S))])
  a3 <- as.numeric(Q1S[i,grep("2019-07-06", colnames(Q1S)):grep("2019-09-28", colnames(Q1S))])
  ax <- as.numeric(Q1S[i,grep("2019-10-05", colnames(Q1S)):grep("2019-12-28", colnames(Q1S))])
  
  ok <- rbind.data.frame(a1,a2,a3)
  rownames(ok) <- c("Q1","Q2","Q3")
  colnames(ok) <- c(1:13)
  
  q1 <- qcc(ok, type="R", nsigmas=3, plot = FALSE)
  
  #2020
  a4 <- as.numeric(Q1S[i,grep("2020-01-04", colnames(Q1S)):grep("2020-03-28", colnames(Q1S))])
  a5 <- as.numeric(Q1S[i,grep("2020-04-04", colnames(Q1S)):grep("2020-06-27", colnames(Q1S))])
  a6 <- as.numeric(Q1S[i,grep("2020-07-04", colnames(Q1S)):grep("2020-10-03", colnames(Q1S))])
  a7 <- as.numeric(Q1S[i,grep("2020-10-10", colnames(Q1S)):grep("2020-12-26", colnames(Q1S))])
  
  
  ok <- rbind.data.frame(a4,a5,a6)
  rownames(ok) <- c("Q1","Q2","Q3")
  colnames(ok) <- c(1:14)
  
  q2 <- qcc(ok, type="R", nsigmas=3, plot = FALSE)
  
  
  
  #Q4 detection
  ok <- rbind.data.frame(a1,a2,a3, ax)
  rownames(ok) <- c("Q1","Q2","Q3","Q4")
  colnames(ok) <- c(1:13)
  
  q4 <- qcc(ok, type="R", nsigmas=9, plot = FALSE)  
  
  ok <- rbind.data.frame(a4,a5,a6, a7)
  rownames(ok) <- c("Q1","Q2","Q3","Q4")
  colnames(ok) <- c(1:14)
  
  q4x <- qcc(ok, type="R", nsigmas=9, plot = FALSE)  
  
  
  
  #Detection of Q1 seasonality ( 2020 + 2019)
  
  if (  (as.numeric(unlist(q2$statistics[1])) > as.numeric(unlist(q2$limits[2]))) && (as.numeric(unlist(q1$statistics[1])) > as.numeric(unlist(q1$limits[2]))) ){
    seas_vec_Q1[i] <- i 
  }
  
  if (  (as.numeric(unlist(q2$statistics[2])) > as.numeric(unlist(q2$limits[2]))) && (as.numeric(unlist(q1$statistics[2])) > as.numeric(unlist(q1$limits[2]))) ){
    seas_vec_Q2[i] <- i 
  } 
  
  if (  (as.numeric(unlist(q2$statistics[3])) > as.numeric(unlist(q2$limits[2]))) && (as.numeric(unlist(q1$statistics[3])) > as.numeric(unlist(q1$limits[2]))) ){
    seas_vec_Q3[i] <- i 
  }  
  
  if (  (as.numeric(unlist(q4$statistics[4])) > as.numeric(unlist(q4$limits[2]))) && (as.numeric(unlist(q4x$statistics[4])) > as.numeric(unlist(q4x$limits[2]))) ){
    seas_vec_Q4[i] <- i 
  }    
  
  
}

seas_vec_Q1 <- seas_vec_Q1[!is.na(seas_vec_Q1)]
seas_vec_Q2 <- seas_vec_Q2[!is.na(seas_vec_Q2)]
seas_vec_Q3 <- seas_vec_Q3[!is.na(seas_vec_Q3)]
seas_vec_Q4 <- seas_vec_Q4[!is.na(seas_vec_Q4)]


Q1SO <- Q1S[seas_vec_Q1,]
Q2SO <- Q1S[seas_vec_Q2,]
Q3SO <- Q1S[seas_vec_Q3,]
Q4SO <- Q1S[seas_vec_Q4,]


#Q1SO <- Q1SO %>% filter(!str_detect(title, "Eyewitness"))


Q1_iso <- cbind.data.frame(Q1SO[,1:3], Q1SO[, grep("2019-01-05", colnames(Q1S)):grep("2019-03-30", colnames(Q1S)) ], Q1_2019_Total = rowSums(Q1SO[,grep("2019-01-05", colnames(Q1S)):grep("2019-03-30", colnames(Q1S))]), 
                           Q1SO[,grep("2020-01-04", colnames(Q1S)):grep("2020-03-28", colnames(Q1S))], Q1_2020_Total = rowSums(Q1SO[,grep("2020-01-04", colnames(Q1S)):grep("2020-03-28", colnames(Q1S)) ]))
Q1_iso <- subset(Q1_iso, Q1_iso$Q1_2019_Total + Q1_iso$Q1_2020_Total >= 50)



Q2_iso <- cbind.data.frame(Q2SO[,1:3], Q2SO[,grep("2019-04-06", colnames(Q1S)):grep("2019-06-29", colnames(Q1S))], Q2_2019_Total = rowSums(Q2SO[,grep("2019-04-06", colnames(Q1S)):grep("2019-06-29", colnames(Q1S))]), 
                           Q2SO[,grep("2020-04-04", colnames(Q1S)):grep("2020-06-27", colnames(Q1S))], Q2_2020_Total = rowSums(Q2SO[,grep("2020-04-04", colnames(Q1S)):grep("2020-06-27", colnames(Q1S)) ]))
Q2_iso <- subset(Q2_iso, Q2_iso$Q2_2019_Total + Q2_iso$Q2_2020_Total >= 50)



Q3_iso <- cbind.data.frame(Q3SO[,1:3], Q3SO[,grep("2019-07-06", colnames(Q1S)):grep("2019-09-28", colnames(Q1S))], Q3_2019_Total = rowSums(Q3SO[,grep("2019-07-06", colnames(Q1S)):grep("2019-09-28", colnames(Q1S))]), 
                           Q3SO[,grep("2020-07-04", colnames(Q1S)):grep("2020-10-03", colnames(Q1S))], Q3_2020_Total = rowSums(Q3SO[,grep("2020-07-04", colnames(Q1S)):grep("2020-10-03", colnames(Q1S))]))
Q3_iso <- subset(Q3_iso, Q3_iso$Q3_2019_Total + Q3_iso$Q3_2020_Total >= 50)



Q4_iso <- cbind.data.frame(Q4SO[,1:3], Q4SO[,grep("2019-10-05", colnames(Q1S)):grep("2019-12-28", colnames(Q1S))], Q4_2019_Total = rowSums(Q4SO[,grep("2019-10-05", colnames(Q1S)):grep("2019-12-28", colnames(Q1S))]), 
                           Q4SO[,grep("2020-10-10", colnames(Q1S)):grep("2020-12-26", colnames(Q1S))], Q4_2020_Total = rowSums(Q4SO[,grep("2020-10-10", colnames(Q1S)):grep("2020-12-26", colnames(Q1S))]))
Q4_iso <- subset(Q4_iso, Q4_iso$Q4_2019_Total + Q4_iso$Q4_2020_Total >= 50)




Q1_iso$Season <- "Q1"
Q2_iso$Season <- "Q2"
Q3_iso$Season <- "Q3"
Q4_iso$Season <- "Q4"


#Create data frame with all Seasonal titles
Q_iso <- rbind.data.frame(Q1_iso[,c(1:3, grep("Season", colnames(Q1_iso)) )], 
                          Q2_iso[,c(1:3, grep("Season", colnames(Q2_iso)) )],
                          Q3_iso[,c(1:3, grep("Season", colnames(Q3_iso)) )],
                          Q4_iso[,c(1:3, grep("Season", colnames(Q4_iso)) )])


#Adding Manual Titles (actual values are added after)

Q_iso <- rbind.data.frame(Q_iso, c("0241283477","9780241283479","Baby Touch and Feel I Love You","Q1")  )
Q_iso <- rbind.data.frame(Q_iso, c("140536288X","9781405362887","Pop-Up Peekaboo! Farm","Q2")  )
Q_iso <- rbind.data.frame(Q_iso, c("0241459001","9780241459003","Baby's First Easter","Q2")  )
Q_iso <- rbind.data.frame(Q_iso, c("1409345610","9781409345619","Kama Sutra A Position A Day","Q1")  )
Q_iso <- rbind.data.frame(Q_iso, c("0241459001","9780241459003","Baby's First Easter","Q2")  )


saveRDS(Q_iso, "Q_iso_UK.csv")