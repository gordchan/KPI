# Starting script for KPI report ----------

# Source required scripts

source("scripts/KPI_report.R")
source("scripts/std_filename.R")

source("scripts/kpi_dates.R")
source("scripts/kpi_filechk.R")
source("scripts/kpi_chkperiod.R")

# Input reporting year & month ----------

y <- 2016
m <- 2

py <- y-1

KPI_dates(y, m)

# Standardise Filenames ----------

dup_queue(y, m)

std_filename(y, m)
std_filename(py, m)

# Source file completness ----------

KPI_filechk(y, m)

# Check source file sample period ----------

KPI_chkperiod(Dates[1,], CurrentYear = TRUE)
KPI_chkperiod(Dates[3,], CurrentYear = FALSE)

# Main Process ----------

KPI_report(y, m)

