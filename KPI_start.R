# Starting script for KPI report

# Source required scripts

source("scripts/KPI_report.R")
source("scripts/std_filename.R")

# Input reporting year & month

y <- 2016
m <- 1

py <- y-1

# Standardise Filenames

dup_queue(y, m)

std_filename(y, m)
std_filename(py, m)

# Main Process

KPI_report(y, m)
