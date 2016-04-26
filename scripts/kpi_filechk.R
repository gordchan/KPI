KPI_filechk <- function(y = 2015, m = 11){
    
    # Validate function input -------------------------------------------------------
    
    if(nchar(y)!=4 | !is.numeric(y)){
        return("Please input a year in 4 digits")
    } else if(!is.numeric(m)|nchar(m)>2){
        return("Please input a month in 1 or 2 integers")
    }
    
    # Get required info -------------------------------------------------------------
    
    # Get the dates
    KPI_dates(y, m)
    
    # Get file paths
    
    path.Current <- file.path("source", Dates[1,])
    path.Previous <- file.path("source", Dates[3,])
    
    # Validations -------------------------------------------------------------------
    
    FAULT <- FALSE # Overall fault detected
    
    # Check if source folders exist
    
    fault_1 <- FALSE
    
    if(!file.exists(path.Current)){
        FAULT <- TRUE
        fault_1 <- TRUE
        warning("ERROR: Folder ", Dates[1,], " not uploaded.")
    }
    if(!file.exists(path.Previous)){
        FAULT <- TRUE
        fault_1 <- TRUE
        warning("ERROR: Folder ", Dates[3,], " not uploaded.")
    }
    if(fault_1==FALSE){
        message("CHECKED: Source folders")
    }else{
        return("ERROR: Source folder not uploaded")
    }
    
    # Check if source files exist
    
    fault_2 <- FALSE
    fault_3 <- FALSE
    
    filelist.Current <- list.files(file.path("source", Dates[1,]), pattern = ".*\\.(xls.?|html)")
    filelist.Previous <- list.files(file.path("source", Dates[3,]), pattern = ".*\\.(xls.?|html)")
    
    ## List of kpi files to check
    
    regx_list <- c("kpi.1 ", "kpi.2 ", "kpi.2.HA ", "kpi.3.1.1 ", "kpi.3.1.2", "kpi.3.2.1", "kpi.3.2.2", "kpi.3.3",
                   "kpi.4.1 ", "kpi.4.2 ", "kpi.4.3", "kpi.5 ", "kpi.6 ", "kpi.6.HA ", "kpi.7 ", "kpi.8 ", "kpi.9 ",
                   "kpi.10 ", "kpi.11.1 ", "kpi.11.1.HA ", "kpi.11.2", "kpi.11.3 ", 
                   "tre.1 ", "tre.2 ", "tre.2.HA ", "tre.2.1 ", "tre.2.1.HA ", "tre.3.1.1 ", "tre.3.1.2 ", "tre.3.2.1 ", "tre.3.2.2 ",
                   "tre.5 ", "tre.6 ", "tre.7 ", "tre.8 ", "tre.8.HA ", "tre.9 ", "tre.10 ", "tre.12 " )
    
    ## Current Year
    
    for(i in 1:length(regx_list)){
        if(sum(grepl(regx_list[i], filelist.Current))==0){
            FAULT <- TRUE
            fault_2 <- TRUE
            warning("ERROR: KPI source file ", regx_list[i], "in ", Dates[1,], " not uploaded.")
        }
    }

    if(fault_2==FALSE){
        message("CHECKED: Source files of Current Year")
    }else{
        return("ERROR: Missing file(s) in Current Year folder")
    }
    
    ## Previous Year
    
    for(i in 1:20){
        if(sum(grepl(regx_list[i], filelist.Previous))==0){
            FAULT <- TRUE
            fault_3 <- TRUE
            warning("ERROR: KPI source file ", regx_list[i], "in ", Dates[3,], " not uploaded.")
        }
    }
    
    if(fault_3==FALSE){
        message("CHECKED: Source files of Previous Year")
    }else{
        return("ERROR: Missing file(s) in Previous Year folder")
    }
    
    # Successful Validation
    
    if(FAULT==FALSE){
        return("CHECKED: ALL CHECKED AND READY TO GO")
    }
}