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



