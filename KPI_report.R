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
source("process_kpi.R")

# Reporting period --------------------------------------------------------

to.YY <- 15 ## Input from func
    from.YY <- to.YY - 1

to.month <- 5 ## Input from func
    to.Mmm <- month(to.month, label = TRUE, abbr = TRUE)
    from.Mnn <- month(to.month + 1, label = TRUE, abbr = TRUE)

to.MmmYY <- paste(to.Mmm , to.YY, sep = "")
    to.MmmYY_ <- paste(" ", to.MmmYY, sep = "") # Keep leading zero for use as text in Excel
    prev.MmmYY <- paste(to.Mmm, from.YY, sep = "")
    
from.MmmYY <- paste(from.Mnn, from.YY, sep = "")

report.period <- paste(from.MmmYY, to.MmmYY, sep = "-")

# Create filepath table ----------------------------------------------------------

KPI_files <- 
    data.frame(
        temp.names = list.files("template", pattern = "(\\.xls$)", full.names = FALSE),
        temp.paths = list.files("template", pattern = "(\\.xls$)", full.names = TRUE), 
    stringsAsFactors = FALSE)

KPI_files <- KPI_files %>%
    mutate(file.paths = gsub("(\\.xls$)", paste(to.MmmYY_, ".xls", sep = ""), temp.paths)) %>%
        mutate(file.paths = gsub("(^template)", "report", file.paths)) %>%
            arrange(desc(temp.names))

# Copy template for use as blank report -------------------------------------------

for (i in 1:nrow(KPI_files)){
    file.copy(from = KPI_files$temp.paths[i],
              to = KPI_files$file.paths[i],
              overwrite = TRUE,
              copy.mode = TRUE, copy.date = FALSE)
}

# Loading xls template ------------------------------------------------------

as.KPI <- loadWorkbook(KPI_files$temp.paths[1])
    sheets.KPI <- getSheets(as.KPI)

# as.SOP <- loadWorkbook(KPI_files$temp.paths[2])
#     sheets.SOP <- getSheets(as.SOP)
#     
# as.TRE <- loadWorkbook(KPI_files$temp.paths[3])
#     sheets.TRE <- getSheets(as.TRE)


# Load and prepare KPI source data -----------------------------------

    # kpi.1
        kpi.1.c <- kpi.1(to.MmmYY)
        kpi.1.t <- kpi.t("kpi.1")
        kpi.1.p <- kpi.1(prev.MmmYY)
    # kpi.2
        kpi.2.t <- kpi.t("kpi.2")
    # kpi.3
        kpi.3.t <- kpi.t("kpi.3")
    # kpi.4
        kpi.4.t <- kpi.t("kpi.4")
    # kpi.5
        kpi.5.t <- kpi.t("kpi.5")
    # kpi.6
        kpi.6.t <- kpi.t("kpi.6")
    # kpi.7
        kpi.7.t <- kpi.t("kpi.7")

# Update reporting month and period -----------------------------------

reporting.dates <- CellBlock(sheets.KPI$source, 1, 2, 2, 1)

    CB.setMatrixData(reporting.dates, as.matrix(c(to.MmmYY_, report.period)), 1, 1)

# Update A&E wait time -----------------------------------------------
    # kpi.1
        addDataFrame(kpi.1.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=3)
        addDataFrame(kpi.1.t, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=13)
        addDataFrame(kpi.1.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=22)
    # kpi.2
        # addDataFrame(kpi.2.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=7, startColumn=3)
        addDataFrame(kpi.2.t, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=7, startColumn=13)
        # addDataFrame(kpi.2.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=7, startColumn=22)
    # kpi.3
        # addDataFrame(kpi.3.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=12, startColumn=3)
        addDataFrame(kpi.3.t, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=12, startColumn=13)
        # addDataFrame(kpi.3.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=12, startColumn=22)
    # kpi.4
        # addDataFrame(kpi.4.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=16, startColumn=3)
        addDataFrame(kpi.4.t, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=16, startColumn=13)
        # addDataFrame(kpi.4.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=16, startColumn=22)
    # kpi.5
        # addDataFrame(kpi.5.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=19, startColumn=3)
        addDataFrame(kpi.5.t, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=19, startColumn=13)
        # addDataFrame(kpi.5.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=19, startColumn=22)
    # kpi.6
        # addDataFrame(kpi.6.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=22, startColumn=3)
        addDataFrame(kpi.6.t, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=22, startColumn=13)
        # addDataFrame(kpi.6.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=22, startColumn=22)
    # kpi.7
        # addDataFrame(kpi.7.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=24, startColumn=3)
        addDataFrame(kpi.7.t, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=24, startColumn=13)
        # addDataFrame(kpi.7.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=24, startColumn=22)

# Saving xls reports ----------------------------------------------------

as.KPI$setForceFormulaRecalculation(TRUE)
# as.SOP$setForceFormulaRecalculation(TRUE)
# as.TRE$setForceFormulaRecalculation(TRUE)

    saveWorkbook(as.KPI, KPI_files$file.paths[1])
#     saveWorkbook(as.SOP, KPI_files$file.paths[2])
#     saveWorkbook(as.TRE, KPI_files$file.paths[3])
    