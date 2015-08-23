
# Libraries ---------------------------------------------------------------

require("xlsx")

source("read_all_sheets.R")

# Filepath table ----------------------------------------------------------

report.month <- "May15"

    filenames <- rbind(
        paste("KWC Clinical KPI ", report.month, ".xls", sep = ""),
        paste("Clinical KPI SOP WT ", report.month, ".xls", sep = ""),
        paste("Clinical KPI Trend ", report.month, ".xls", sep = "")
    )
    
    filepaths <- rbind(
        file.path("template", filenames[1]),
        file.path("template", filenames[2]),
        file.path("template", filenames[3])
    )

KPI_files <- data.frame(filenames, filepaths, stringsAsFactors = FALSE)

# Loading xls report ------------------------------------------------------

as.KPI <- loadWorkbook(KPI_files$filepaths[1])
    sheets.KPI <- getSheets(as.KPI)

# as.SOP <- loadWorkbook(KPI_files$filepaths[2])
#     sheets.KPI <- getSheets(as.SOP)
#     
# as.TRE <- loadWorkbook(KPI_files$filepaths[3])
#     sheets.KPI <- getSheets(as.TRE)


# Loading KPI source data -------------------------------------------------


test_value.c <- matrix(seq(0.01, 0.40, 0.01), nrow = 4, ncol = 10, byrow = TRUE)
test_value.t <- matrix(seq(0.01, 0.40, 0.01), nrow = 4, ncol = 10, byrow = TRUE)
test_value.p <- matrix(seq(0.01, 0.36, 0.01), nrow = 4, ncol = 10, byrow = TRUE)

AE_WT <- read_range("source/AE WT May15.xlsx", 5:12, 1:14)

    
# Define CellBlocks -------------------------------------------------------

AE_KWC.c <- CellBlock(sheets.KPI$source, 3, 3, 4, 10)
AE_KWC.t <- CellBlock(sheets.KPI$source, 3, 13, 4, 10)
AE_KWC.p <- CellBlock(sheets.KPI$source, 3, 32, 4, 10)

    CB.setMatrixData(AE_KWC.c, test_value.c, 1, 1)
    CB.setMatrixData(AE_KWC.t, test_value.t, 1, 1)
    CB.setMatrixData(AE_KWC.p, test_value.p, 1, 1)

            saveWorkbook(as.KPI, KPI_files$filepaths[1])