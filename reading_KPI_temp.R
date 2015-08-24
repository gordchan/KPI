
# I. Libraries ---------------------------------------------------------------

require("xlsx")
require("dplyr")

source("read_xlsx.R")


# II. Create filepath table ----------------------------------------------------------

report.month <- " May15" ## Keep leading zero to keep as text in excel
report.period <- "Jun14-May15"

    filenames <- rbind(
        paste("KWC Clinical KPI", report.month, ".xls", sep = ""),
        paste("Clinical KPI SOP WT", report.month, ".xls", sep = ""),
        paste("Clinical KPI Trend", report.month, ".xls", sep = "")
    )
    
    filepaths <- rbind(
        file.path("template", filenames[1]),
        file.path("template", filenames[2]),
        file.path("template", filenames[3])
    )

KPI_files <- data.frame(filenames, filepaths, stringsAsFactors = FALSE)

# III. Loading xls report ------------------------------------------------------

as.KPI <- loadWorkbook(KPI_files$filepaths[1])
    sheets.KPI <- getSheets(as.KPI)

# as.SOP <- loadWorkbook(KPI_files$filepaths[2])
#     sheets.SOP <- getSheets(as.SOP)
#     
# as.TRE <- loadWorkbook(KPI_files$filepaths[3])
#     sheets.TRE <- getSheets(as.TRE)


# IV. Loading KPI source data -------------------------------------------------

    source("process_kpi.R")
    
# Update reporting month and period ---------------------------------------

reporting <- CellBlock(sheets.KPI$source, 1, 2, 2, 1)

    CB.setMatrixData(reporting, as.matrix(c(report.month, report.period)), 1, 1)

# Update A&E wait time ----------------------------------------------------

# AE_WT.cb.c <- CellBlock(sheets.KPI$source, 3, 3, 4, 10)
# AE_WT.cb.t <- CellBlock(sheets.KPI$source, 3, 13, 4, 10)
# AE_WT.cb.p <- CellBlock(sheets.KPI$source, 3, 32, 4, 10)
# 
#     CB.setMatrixData(AE_WT.cb.c, AE_WT.prod, 1, 1)
#     CB.setMatrixData(AE_WT.cb.t, AE_WT.prod, 1, 1)
#     CB.setMatrixData(AE_WT.cb.p, AE_WT.prod, 1, 1)
    
addDataFrame(AE_WT.prod, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=3)
addDataFrame(AE_WT.prod, sheets.KPI$source, col.names=FALSE, row.names=FALSE,  startRow=3, startColumn=13)
addDataFrame(AE_WT.prod, sheets.KPI$source, col.names=FALSE, row.names=FALSE,  startRow=3, startColumn=23)

# Saving xls reports ----------------------------------------------------

    saveWorkbook(as.KPI, KPI_files$filepaths[1])