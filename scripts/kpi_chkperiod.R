KPI_chkperiod_fx <- function(Mmm, regx, type, CurrentYear){
    
    if(!exists("raw_range", mode="function")) source("scripts/read_xlsx.R")
    if(!exists("kpi_source_helper", mode="function")) source("scripts/kpi_helper.R")

    require(lubridate)
    
# Validate function input -------------------------------------------------------

    if(!(type=="CDARS" | type=="EIS" | type=="MIPO" | type=="CDARS_Std" | type=="RAD")){
        return("Please input a type of CDARS/EIS/MIPO/CDARS_Std/RAD only")
    }
    
    if(!exists("Dates")){
        return("Please run KPI_dates first")
    }
    
# Get required info -------------------------------------------------------------

    start <- dmy("1-1-1987")
    end <- dmy("1-1-1987")
    
    
    kpi_source_helper(Mmm)
    
    path <- grep(regx, source.kpi$filepath, value = TRUE)
    
    if(CurrentYear==TRUE){
        kpi_end <- dmy(Dates[7,])
    }else{
        kpi_end <- dmy(Dates[8,])
    }
    
    
    kpi_start <- kpi_end
        day(kpi_start) <- day(kpi_start) +1
            year(kpi_start) <- year(kpi_start) -1
    
# Get file paths & lists

    if(type=="CDARS"){
        # df <- raw_range(input = path, ri = 1:100, ci = 1, si = 1)
        lines <- readLines(con = path, n = 7)
        lines <- lines[7]
    }else if(type=="EIS"){
        df <- raw_range(input = path, ri = 1:100, ci = 2, si = 1)
    }else if(type=="MIPO"){
        df <- raw_range(input = path, ri = 1:50, ci = 1, si = 2)
    }else if(type=="CDARS_Std"){
        # df <- raw_range(input = path, ri = 1:4, ci = 2, si = 1)
        lines <- readLines(con = path, n = 66)
        lines <- lines[66]
    }else if(type=="RAD"){
        df <- raw_range(input = path, ri = 2:120, ci = 2, si = 1)
    }

    
# Validations -------------------------------------------------------------------
    
    # Raw info to be validated
    
    if(type=="CDARS"){
        # cell <- df[which(grepl("(Date between .*[[:digit:]]{4})", df[,1])),1] # Single cell containing info
        cell <- lines
    }else if(type=="EIS"){
        cell <- df[which(grepl("(Month:.*[[:digit:]]{2})", df[,1])),1] # Single cell containing info
    }else if(type=="MIPO"){
        cell_start <- which(df=="Period")+1
        cell <- data.frame(df[cell_start:nrow(df), 1]) # Multiple cells containing the period
        names(cell) <- "LabelPeriod"
        
        cell <- cell %>% unique() %>% mutate(From = gsub("(.*?) - .*", "\\1", LabelPeriod), To = gsub(".*? - (.*)", "\\1", LabelPeriod)) %>%
            mutate(FromM = gsub("([a-zA-Z]{3}) ([0-9]{4})", "\\1", From), ToM = gsub("([a-zA-Z]{3}) ([0-9]{4})", "\\1", To)) %>%
            mutate(FromM = match(FromM, month.abb), ToM = match(ToM, month.abb)) %>%
            mutate(FromY = gsub("([a-zA-Z]{3}) ([0-9]{4})", "\\2", From), ToY = gsub("([a-zA-Z]{3}) ([0-9]{4})", "\\2", To)) %>%
            mutate(From = paste("1", FromM, FromY, sep = "-"), To = paste("1", ToM, ToY, sep = "-"))
        
    }else if(type=="CDARS_Std"){
        # cell <- df[1,1]
        cell <- lines
    }else if(type=="RAD"){
        cell <- df
        names(cell) <- "LabelPeriod"
        
        cell <- cell %>% unique() %>% mutate(From = gsub("(.*?) - .*", "\\1", LabelPeriod), To = gsub(".*? - (.*)", "\\1", LabelPeriod)) %>%
            mutate(FromM = gsub("([a-zA-Z]{3}) ([0-9]{4})", "\\1", From), ToM = gsub("([a-zA-Z]{3}) ([0-9]{4})", "\\1", To)) %>%
            mutate(FromM = match(FromM, month.abb), ToM = match(ToM, month.abb)) %>%
            mutate(FromY = gsub("([a-zA-Z]{3}) ([0-9]{4})", "\\2", From), ToY = gsub("([a-zA-Z]{3}) ([0-9]{4})", "\\2", To)) %>%
            mutate(From = paste("1", FromM, FromY, sep = "-"), To = paste("1", ToM, ToY, sep = "-"))
    }

    # Parse info to date format
    
    if(type=="CDARS"){
        start <- sub(".*?between ([0-9]{2}.[0-9]{2}.[0-9]{4}) and.*", "\\1", cell)
        end <- sub(".*?Date between .* and ([0-9]{2}.[0-9]{2}.[0-9]{4}).*", "\\1", cell)
        
        start <- dmy(start)
        end <- dmy(end)
        
    }else if(type=="EIS"){
        start <- sub("^Month:(.*?) to .*", "\\1", cell)
        end <- sub(".*?to (.*)", "\\1", cell)
        
        startM <- sub("([a-zA-Z]{3}) ([0-9]{2})", "\\1", start)
        startM <- match(startM, month.abb)
        startY <- sub("([a-zA-Z]{3}) ([0-9]{2})", "\\2", start)
        
        endM <- sub("([a-zA-Z]{3}) ([0-9]{2})", "\\1", end)
        endM <- match(endM, month.abb)
        endY <- sub("([a-zA-Z]{3}) ([0-9]{2})", "\\2", end)
        
        start <- paste("1", startM, startY)
        end <- paste("1", endM, endY)
        
        start <- dmy(start)
        end <- dmy(end)
            month(end) <- month(end) +1
            day(end) <- day(end) -1
        
    }else if(type=="MIPO" | type=="RAD"){
        cell <- cell %>% mutate(start = dmy(From), end = dmy(To))
        month(cell$end) <- month(cell$end) +1
        day(cell$end) <- day(cell$end) -1
        
        row_i <- which(cell$end==kpi_end) # Check which row is relevant to input reporting period
        
        end <- cell$end[row_i]
        start <- cell$start[row_i]
        
    }else if(type=="CDARS_Std"){
        start <- sub(".*From ([0-9]{2}-[a-zA-Z]{3}-[0-9]{4}) to .*", "\\1", cell)
        end <- sub(".*to ([0-9]{2}-[a-zA-Z]{3}-[0-9]{4}) .*", "\\1", cell)
        
        startD <- sub("([0-9]{2})-([a-zA-Z]{3})-([0-9]{4})", "\\1", start)
        endD <- sub("([0-9]{2})-([a-zA-Z]{3})-([0-9]{4})", "\\1", end)
        
        startM <- sub("([0-9]{2})-([a-zA-Z]{3})-([0-9]{4})", "\\2", start)
        endM <- sub("([0-9]{2})-([a-zA-Z]{3})-([0-9]{4})", "\\2", end)
        
        startM <- match(startM, month.abb)
        endM <- match(endM, month.abb)
        
        startY <- sub("([0-9]{2})-([a-zA-Z]{3})-([0-9]{4})", "\\3", start)
        endY <- sub("([0-9]{2})-([a-zA-Z]{3})-([0-9]{4})", "\\3", end)
        
        start <- paste(startD, startM, startY, sep = "-")
        end <- paste(endD, endM, endY, sep = "-")
        
        start <- dmy(start)
        end <- dmy(end)
    }
            
    # Check if dates are same

    if(start!=kpi_start){
        warning("ERROR: ", regx, "in ", Mmm, " , period start is incorrect")
    }else if(end!=kpi_end){
        warning("ERROR: ", regx, "in ", Mmm, " , period end is incorrect")
    }else{
        message(paste("CHECKED: Period of ", regx, "in ", Mmm, " is correct", sep = ""))
    }
    
}



KPI_chkperiod <- function(Mmm, CurrentYear){
    
    # Check sampling period of source files

    KPI_chkperiod_fx(Mmm, regx = "kpi.1 ", type = "EIS", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.2 ", type = "EIS", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.2.HA ", type = "EIS", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.3.1.1 ", type = "CDARS", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.3.1.2 ", type = "CDARS", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.3.2.1 ", type = "CDARS", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.3.2.2 ", type = "CDARS", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.3.3 ", type = "CDARS_Std", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.4.1 ", type = "CDARS", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.4.2 ", type = "CDARS", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.4.3 ", type = "MIPO", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.5 ", type = "EIS", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.6 ", type = "MIPO", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.6.HA ", type = "MIPO", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.7 ", type = "RAD", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.8 ", type = "CDARS", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.9 ", type = "EIS", CurrentYear)
    KPI_chkperiod_fx(Mmm, regx = "kpi.10 ", type = "EIS", CurrentYear)

    
    if(CurrentYear==TRUE){
        #
        KPI_chkperiod_fx(Mmm, regx = "kpi.11.1 ", type = "EIS", CurrentYear)
        KPI_chkperiod_fx(Mmm, regx = "kpi.11.1.HA ", type = "EIS", CurrentYear)
        #
        KPI_chkperiod_fx(Mmm, regx = "kpi.11.2 ", type = "EIS", CurrentYear)
        KPI_chkperiod_fx(Mmm, regx = "kpi.11.3 ", type = "EIS", CurrentYear)
    }

    
    # Successful Validation
    
    return("CHECKED: ALL PERIOD CHECKED AND READY TO GO")
    
    
}