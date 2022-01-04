library.path <- .libPaths("C:/Users/steph/Documents/R/win-library/4.0")
cat("Dashboard update start\n")

suppressMessages(library(tidyverse, lib.loc = library.path))
suppressMessages(library(httr, lib.loc = library.path))

r <- GET("https://6g6ddjwzc7.execute-api.eu-west-1.amazonaws.com/prod/refreshMVWGeneral?dataSource=mvw_amazon_forecast_inventory&markerColumnName=uid&limit=20000&index=amz_forecasting_inventory",
         add_headers( .headers =   c( 'x-api-key'= 'IEQP4Z17jq6LzrS3Veoww5ThmodRwiMJ1vRqmAbb',
                          'index'= 'amz_forecasting_inventory',
                          'Access-Control-Allow-Origin'= '*',
                          'Access-Control-Allow-Credentials'= 'true',
                          'type'= 'text' ) )
         )

warn_for_status(r)





