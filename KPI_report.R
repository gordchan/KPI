# Reading KPI Template
#
# To load Excel KPI template and write in consolidated KPI data.
#
# Gordon CHAN
#
# Aug 2015


# y = 2015
# m = 5

kpi_report <- function(y = 2015, m = 5){

# Libraries -----------------------------------------------------------------------

require("xlsx")
require("dplyr")
require("lubridate")

source("read_xlsx.R")
source("process_kpi.R")


# Validate function input -------------------------------------------------

    if(nchar(y)!=4 | !is.numeric(y)){
        return("Please input a year in 4 digits")
    } else if(!is.numeric(m)){
        return("Please input a month in integers")
    }

# Reporting period --------------------------------------------------------

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

dates <- data.frame(dates = c(to.MmmYY_,
                    cy.period,
                    py.period),
                    stringsAsFactors = FALSE)

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
        kpi.1.p <- kpi.1(prev.to.MmmYY)
    # kpi.2
        kpi.2.c <- kpi.2(to.MmmYY)
        kpi.2.p <- kpi.2(prev.to.MmmYY)
    # kpi.3
        # kpi.3.1.c <- kpi.10(to.MmmYY)
        # kpi.3.1.p <- kpi.10(prev.to.MmmYY)
        # kpi.3.2.c <- kpi.10(to.MmmYY)
        # kpi.3.2.p <- kpi.10(prev.to.MmmYY)
        # kpi.3.3.c <- kpi.10(to.MmmYY)
        # kpi.3.3.p <- kpi.10(prev.to.MmmYY)
        kpi.3.4.c <- kpi.10(to.MmmYY)
        kpi.3.4.p <- kpi.10(prev.to.MmmYY)
    # kpi.4
        kpi.4.1.c <- kpi.4.1(to.MmmYY)
        kpi.4.1.p <- kpi.4.1(prev.to.MmmYY)
        kpi.4.2.c <- kpi.4.2(to.MmmYY)
        kpi.4.2.p <- kpi.4.2(prev.to.MmmYY)
        # kpi.4.3.c <- kpi.4.3(to.MmmYY)
        # kpi.4.3.p <- kpi.4.3(prev.to.MmmYY)
    # kpi.5
        # kpi.5.1.c <- kpi.5(to.MmmYY)
        # kpi.5.1.p <- kpi.5(prev.to.MmmYY)
        kpi.5.2.c <- kpi.5("kpi.5.2", to.MmmYY)
        kpi.5.2.p <- kpi.5("kpi.5.2", prev.to.MmmYY)
        kpi.5.3.c <- kpi.5("kpi.5.3", to.MmmYY)
        kpi.5.3.p <- kpi.5("kpi.5.3", prev.to.MmmYY)
    # kpi.6
        # kpi.6.c <- kpi.6(to.MmmYY)
        # kpi.6.p <- kpi.6(prev.to.MmmYY)
    # kpi.7
        # kpi.7.c <- kpi.7(to.MmmYY)
        # kpi.7.p <- kpi.7(prev.to.MmmYY)

# Update reporting month and period -----------------------------------

    addDataFrame(dates, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=1, startColumn=2)

# Write to Excel templates -----------------------------------------------
    # kpi.t ## Prototype not functioning pending rewrite of kpi.t()
            for (i in 1:25){
                kpi.t.i <- kpi.t(i, to.MmmYY)
                addDataFrame(kpi.t.i, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=(3+i), startColumn=13)
            }
    # kpi.1
        addDataFrame(kpi.1.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=4, startColumn=3)
        
        addDataFrame(kpi.1.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=4, startColumn=22)
    # kpi.2
        addDataFrame(kpi.2.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=8, startColumn=3)
        
        addDataFrame(kpi.2.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=8, startColumn=22)
    # kpi.3
        # addDataFrame(kpi.3.1.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=13, startColumn=3)
        # addDataFrame(kpi.3.2.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=14, startColumn=3)
        # addDataFrame(kpi.3.3.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=15, startColumn=3)
        addDataFrame(kpi.3.4.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=16, startColumn=3)
        
        # addDataFrame(kpi.3.1.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=13, startColumn=22)
        # addDataFrame(kpi.3.2.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=14, startColumn=22)
        # addDataFrame(kpi.3.3.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=15, startColumn=22)
        addDataFrame(kpi.3.4.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=16, startColumn=22)
    # kpi.4
        addDataFrame(kpi.4.1.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=17, startColumn=3)
        addDataFrame(kpi.4.2.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=18, startColumn=3)
        # addDataFrame(kpi.4.3.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=19, startColumn=3)

        addDataFrame(kpi.4.1.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=17, startColumn=22)
        addDataFrame(kpi.4.2.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=18, startColumn=22)
        # addDataFrame(kpi.4.3.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=19, startColumn=22)
        
    # kpi.5
        # addDataFrame(kpi.5.1.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=20, startColumn=3)
        addDataFrame(kpi.5.2.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=21, startColumn=3)
        addDataFrame(kpi.5.3.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=22, startColumn=3)
        
        # addDataFrame(kpi.5.1.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=20, startColumn=22)
        addDataFrame(kpi.5.2.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=21, startColumn=22)
        addDataFrame(kpi.5.3.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=22, startColumn=22)
    # kpi.6
        # addDataFrame(kpi.6.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=23, startColumn=3)

        # addDataFrame(kpi.6.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=23, startColumn=22)
    # kpi.7
        # addDataFrame(kpi.7.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=25, startColumn=3)

        # addDataFrame(kpi.7.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=25, startColumn=22)

# Saving xls reports ----------------------------------------------------

as.KPI$setForceFormulaRecalculation(TRUE)
# as.SOP$setForceFormulaRecalculation(TRUE)
# as.TRE$setForceFormulaRecalculation(TRUE)

    saveWorkbook(as.KPI, KPI_files$file.paths[1])
#     saveWorkbook(as.SOP, KPI_files$file.paths[2])
#     saveWorkbook(as.TRE, KPI_files$file.paths[3])
}