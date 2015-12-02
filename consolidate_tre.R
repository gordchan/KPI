#
# Consolidate data for tre report
#

# Load and prepare TRE source data -----------------------------------

# 12 Months Trend

tre.4 <- kpi.3.3(Dates[1,], trend = TRUE)



# Update reporting month and period -----------------------------------

addDataFrame(Series, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=2, startColumn=2)

# Write to Excel templates -----------------------------------------------

# 12 Months Trend

addDataFrame(tre.4[c(3:4),], sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=91, startColumn=2)
