# Read XLSX
#
# Collection of custom functions to handle I/O with excel spreadsheets.
#
# Gordon CHAN
#
# Aug 2015


# read_all_sheets ---------------------------------------------------------

## A function to return all sheets in xlsx to a list with dataframe elements

read_all_sheets <- function (x){
  
  require("xlsx")
  
  input <- x
  
  wb <- loadWorkbook(input)
  sheets <- getSheets(wb)
  sheet_no <- length(sheets)
  
  sheet_index <- c(1:sheet_no)
  
  all_sheets <- lapply(sheet_index, function (x) {read.xlsx(file = input, sheetIndex = x, stringsAsFactors = FALSE)})
  
  names(all_sheets) <- names(sheets)
  
  all_sheets
  
}


# read_range --------------------------------------------------------------

## A function to return a specific range of a sheet from a excel spreadsheets file

read_range <- function (input, ri, ci){
    
    require("xlsx")

    all_range <- read.xlsx(file=input, sheetIndex=1, rowIndex=ri, colIndex=ci,
              as.data.frame=TRUE, header=FALSE, colClasses=NA,
              keepFormulas=FALSE, stringsAsFactors = FALSE)
    
    for (i in 1:length(all_range)){
        if (!is.na(all_range[2,i])){
            
            all_range[1,i] <- all_range[2,i]
        
        }
    }
        if (sum(grepl("Overall", all_range[1,]))==2){
                all_range[1,] <- gsub("(^Overall \\(|\\))", "", all_range[1,])
                all_range[1,] <- gsub("Overall$", "HA", all_range[1,]) ## Will be triggered if KWC is repeated in same axis
        } else {
                all_range[1,] <- gsub("Overall$", "KWC", all_range[1,])
        }

            
            var_col_i <- which(is.na(all_range[1,]))
            
            for (i in 1:length(var_col_i)){
                all_range[1,i] <- paste("var",i, sep = "")
            }
    
    names(all_range) <- all_range[1,]
        all_range <- all_range[-c(1,2),]
    
    all_range$var1 <- gsub("(^ *)", "", all_range$var1)
        # rownames(all_range) <- all_range$var1 ## Will casue error if row names are redundant

    all_range
    
}


# Read raw range ----------------------------------------------------------

raw_range <- function (input, ri, ci){
    
    require("xlsx")
    
    raw_range <- read.xlsx(file=input, sheetIndex=1, rowIndex=ri, colIndex=ci,
                           as.data.frame=TRUE, header=FALSE, colClasses=NA,
                           keepFormulas=FALSE, stringsAsFactors = FALSE)
    raw_range
}