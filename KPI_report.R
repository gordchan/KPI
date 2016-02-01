# Reading KPI Template
#
# To load Excel KPI template and write in consolidated KPI data.
#
# Gordon CHAN
#
# Aug 2015


kpi_report <- function(y = 2015, m = 10){

# Log ---------------------------------------------------------------------

print(paste("Job Started:", Sys.time()))
    
fileConn<-"Log.txt"
cat(paste("Job Started:", Sys.time()), file=fileConn, append=TRUE, sep = "\n")
    
# Libraries ---------------------------------------------------------------

require("xlsx")
require("dplyr")
require("reshape2")
require("lubridate")

# Reporting period --------------------------------------------------------

# if(RAD==TRUE){RAD<<-TRUE}else{RAD<<-FALSE}
    
reportDates <- function(y, m){
    
    require(lubridate)
    
    to.YY <- paste(unlist(strsplit(as.character(y), ""))[3:4], collapse = "") ## Input from func
    from.YY <- paste(unlist(strsplit(as.character(y-1), ""))[3:4], collapse = "") ## Input from func
    prev.from.YY <- paste(unlist(strsplit(as.character(y-2), ""))[3:4], collapse = "") ## Input from func
    
    to.month <- m ## Input from func
    to.Mmm <- month(to.month, label = TRUE, abbr = TRUE)
    from.Mmm <- month(to.month + 1, label = TRUE, abbr = TRUE)
    
    to.MmmYY <- paste(to.Mmm , to.YY, sep = "")
    to.MmmYY_ <- paste(" ", to.MmmYY, sep = "") # Keep leading zero for use as text in Excel
    
    from.MmmYY <- paste(from.Mmm, from.YY, sep = "")
    
    prev.to.MmmYY <- paste(to.Mmm, from.YY, sep = "")
    prev.from.MmmYY <- paste(from.Mmm, prev.from.YY, sep = "")
    
    cy.period <- paste(from.MmmYY, to.MmmYY, sep = "-")
    py.period <- paste(prev.from.MmmYY, prev.to.MmmYY, sep = "-")
    
    eom <- ymd(paste(y, "-", m + 1, "-", 1)) # Find end date of reporting period this year
        day(eom) <- day(eom) -1
    som <- ymd(paste(y-1, "-", m + 1, "-", 1)) # Find start date of reporting period this year
    
    date.period <- paste(day(som), month(som, label = TRUE, abbr = TRUE), year(som), "-", day(eom), month(eom, label = TRUE, abbr = TRUE), year(eom))
    date.eom <- paste(day(eom), "/", month(eom), "/", year(eom), sep = "")
    
    df <- data.frame(dates = c(to.MmmYY, # 1 Reporting month
                               to.MmmYY_, # 2 Reporting month w/ leading zero
                               prev.to.MmmYY, # 3 Last year's reporting month
                               cy.period, # 4 Current year reporting period
                               py.period,# 5 Previous year reporting period
                               date.period, # 6 Reporting period in dates (for Excel report)
                               date.eom, # 7 EOM date (for Excel report)
                               y, # 8 Raw Reporting Year
                               m # 9 Raw Reporting Month
                               ), 
                        stringsAsFactors = FALSE)
}

dateSeries <- function(y, m){
    
    require(lubridate)
    require(dplyr)
    
    eos <- ymd(paste(y, sprintf("%02.f", m), "01", sep = ""))
    
    series <- data.frame(date = .POSIXct(rep(NA, 12)))
    
    for(i in 0:11){
        series[12-i, 1] <- eos - months(i)
    }
    
    series <- series %>% mutate(month = sapply(date, function(x) month(x, label = TRUE))) %>%
        mutate(year = sapply(date, function(x) paste(unlist(strsplit(as.character(year(x)), ""))[3:4], collapse = ""))) %>%
        mutate(label = paste(month, year))
    
    seriesLabel <- series %>% select(label) %>% t()
    
}

Series <<- dateSeries(y = y, m = m)

Dates <<- reportDates(y = y, m = m)

source("read_xlsx.R")
source("process_kpi.R")


# Validation  -------------------------------------------------

# Function input

    if(nchar(y)!=4 | !is.numeric(y)){
        return("Please input a year in 4 digits")
    } else if(!is.numeric(m)){
        return("Please input a month in integers")
    }

# Source file completness




# Create filepath table ----------------------------------------------------------

KPI_files <- 
    data.frame(
        temp.names = list.files("template", pattern = "(\\.xls$)", full.names = FALSE),
        temp.paths = list.files("template", pattern = "(\\.xls$)", full.names = TRUE), 
    stringsAsFactors = FALSE)

KPI_files <- KPI_files %>%
    mutate(file.paths = gsub("(\\.xls$)", paste(Dates[2,], ".xls", sep = ""), temp.paths)) %>%
        mutate(file.paths = gsub("(^template)", file.path("report", Dates[1,]), file.paths)) %>%
            arrange(desc(temp.names))

# Creat directory Eg"Sep15" under report ------------------------------------------

dir.create(file.path("report", Dates[1,]), showWarnings = FALSE)

# Copy template for use as blank report -------------------------------------------

for (i in 1:nrow(KPI_files)){
    file.copy(from = KPI_files$temp.paths[i],
              to = KPI_files$file.paths[i],
              overwrite = TRUE,
              copy.mode = TRUE, copy.date = FALSE)
}

# Loading xls template ------------------------------------------------------

as.KPI <<- loadWorkbook(KPI_files$file.paths[which(grepl("KWC Clinical", KPI_files[,1]))])
    sheets.KPI <<- getSheets(as.KPI)

as.SOP <<- loadWorkbook(KPI_files$file.paths[which(grepl("SOP", KPI_files[,1]))])
    sheets.SOP <<- getSheets(as.SOP)
    
as.TRE <<- loadWorkbook(KPI_files$file.paths[which(grepl("Trend", KPI_files[,1]))])
    sheets.TRE <<- getSheets(as.TRE)


    
# Load and prepare KPI source data -----------------------------------

source("consolidate_kpi.R")

source("consolidate_sop.R")

source("consolidate_tre.R")
    
# Saving xls reports ----------------------------------------------------

as.KPI$setForceFormulaRecalculation(TRUE)
as.SOP$setForceFormulaRecalculation(TRUE)
as.TRE$setForceFormulaRecalculation(TRUE)

    saveWorkbook(as.KPI, KPI_files$file.paths[which(grepl("KWC Clinical", KPI_files[,1]))])
    saveWorkbook(as.SOP, KPI_files$file.paths[which(grepl("SOP", KPI_files[,1]))])
    saveWorkbook(as.TRE, KPI_files$file.paths[which(grepl("Trend", KPI_files[,1]))])
    
# Log ---------------------------------------------------------------------

print(paste("Job End:", Sys.time()))   
    
cat(paste("Job End:", Sys.time()), file=fileConn, append=TRUE, sep = "\n")

    
}

