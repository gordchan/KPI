# Reading KPI Template
#
# To load Excel KPI template and write in consolidated KPI data.
#
# Gordon CHAN
#
# Aug 2015

# Libraries -----------------------------------------------------------------------

require("xlsx")
require("dplyr")
require("lubridate")

source("read_xlsx.R")

# Reporting period --------------------------------------------------------

report.year <- "15"
report.month <- paste(month(5, label = TRUE, abbr = TRUE), report.year, sep = "")
    report.month_ <- paste(" ", report.month, sep = "") ## Keep leading zero to keep as text in excel
prev.month <- paste(month(5-1, label = TRUE, abbr = TRUE), report.year, sep = "")
report.period <- paste(prev.month, report.month, sep = "-")

# Create filepath table ----------------------------------------------------------

KPI_files <- 
    data.frame(
        filenames = list.files("template", pattern = "(\\.xlsx?$)", full.names = FALSE),
        filepaths = list.files("template", pattern = "(\\.xlsx?$)", full.names = TRUE), 
    stringsAsFactors = FALSE)

# Loading xls report ------------------------------------------------------

as.KPI <- loadWorkbook(KPI_files$filepaths[1])
    sheets.KPI <- getSheets(as.KPI)

# as.SOP <- loadWorkbook(KPI_files$filepaths[2])
#     sheets.SOP <- getSheets(as.SOP)
#     
# as.TRE <- loadWorkbook(KPI_files$filepaths[3])
#     sheets.TRE <- getSheets(as.TRE)


# Load and prepare KPI source data -----------------------------------

    source("process_kpi.R")
    
# Update reporting month and period -----------------------------------

reporting <- CellBlock(sheets.KPI$source, 1, 2, 2, 1)

    CB.setMatrixData(reporting, as.matrix(c(report.month, report.period)), 1, 1)

# Update A&E wait time -----------------------------------------------

addDataFrame(AE_WT.prod, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=3)
addDataFrame(AE_WT.prod, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=13)
addDataFrame(AE_WT.prod, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=23)

# Saving xls reports ----------------------------------------------------

    saveWorkbook(as.KPI, KPI_files$filepaths[1])
#     saveWorkbook(as.SOP, KPI_files$filepaths[2])
#     saveWorkbook(as.TRE, KPI_files$filepaths[3])
    