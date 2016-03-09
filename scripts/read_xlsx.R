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

read_range <- function (input, ri, ci, si = 1){
    
    require("xlsx")

    all_range <- read.xlsx(file=input, sheetIndex=si, rowIndex=ri, colIndex=ci,
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

raw_range <- function (input, ri, ci, si = 1){
    
    require("xlsx")
    
    raw_range <- read.xlsx(file=input, sheetIndex=si, rowIndex=ri, colIndex=ci,
                           as.data.frame=TRUE, header=FALSE, colClasses=NA,
                           keepFormulas=FALSE, stringsAsFactors = FALSE)
    raw_range
}


# Read (semi-)fuzzy range --------------------------------------------------------

# Read table of uncertain length given approx. lower bound of rows
#
# with the number of columns assumed known. It searches for the end of table 
#
# assuming the last column would contain only table data, but no text in other cells.

fuzzy_range <- function (input, fuzzy.ri, ci, si = 1){
    
    require("xlsx")
    require("dplyr")
    
    # Read spreadsheet with deliberate over read to accomodate variations in table length
    
    fuzzy_range <- read.xlsx(file=input, sheetIndex=si, rowIndex=fuzzy.ri, colIndex=ci,
                           as.data.frame=TRUE, header=FALSE, colClasses=NA,
                           keepFormulas=FALSE, stringsAsFactors = FALSE)
    
    # Search for end of table
    
    for (i in 1:nrow(fuzzy_range)){
        if (!is.na(fuzzy_range[i,ncol(fuzzy_range)]) & is.na(fuzzy_range[(i+1),ncol(fuzzy_range)])){
            eor <- i
        }
    }
    
    # Cut non-table rows
    
    fuzzy_range <- fuzzy_range[1:eor,]
    
    fuzzy_range
}
