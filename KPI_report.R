# Reading KPI Template
#
# To load Excel KPI template and write in consolidated KPI data.
#
# Gordon CHAN
#
# Aug 2015


kpi_report <- function(y = 2015, m = 10){

# Log ---------------------------------------------------------------------
ptm <- proc.time() # Start timer

print(paste("Job Started:", Sys.time()))
    
fileConn<-"Log.txt"
cat(paste("Job Started:", Sys.time()), file=fileConn, append=TRUE, sep = "\n")
    
# Libraries ---------------------------------------------------------------

require("xlsx")
require("dplyr")
require("reshape2")
require("lubridate")

source("scripts/kpi_dates.R")
source("scripts/kpi_filechk.R")
source("scripts/kpi_chkperiod.R")

source("scripts/read_xlsx.R")
source("scripts/process_kpi.R")

source("scripts/kpi_helper.R")

# Reporting period --------------------------------------------------------

KPI_dates(y, m)


# Validation  -------------------------------------------------

# Function input

if(nchar(y)!=4 | !is.numeric(y)){
    return("Please input a year in 4 digits")
} else if(!is.numeric(m)|nchar(m)>2){
    return("Please input a month in 1 or 2 integers")
}

# Source file completness

KPI_filechk(y, m)

# Check source file sample period

KPI_chkperiod(Dates[1,], CurrentYear = TRUE)
KPI_chkperiod(Dates[3,], CurrentYear = FALSE)

# Successful validation

cat(paste("Validation successful, generating report now", Sys.time()), file=fileConn, append=TRUE, sep = "\n")
print("Validation successful, generating report...")

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

source("scripts/consolidate_kpi.R")

source("scripts/consolidate_sop.R")

source("scripts/consolidate_tre.R")
    
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

print(proc.time() - ptm) # Stop timer

}