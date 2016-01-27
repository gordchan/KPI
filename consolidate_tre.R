#
# Consolidate data for tre report
#

# Load and prepare TRE source data -----------------------------------

# 12 Months Trend

## % AE cases within pledged WT
tre.1.1 <- tre.1(Dates[1,], triage = "Tri 1")
tre.1.2 <- tre.1(Dates[1,], triage = "Tri 2")
tre.1.3 <- tre.1(Dates[1,], triage = "Tri 3")
tre.1.4 <- tre.1(Dates[1,], triage = "Tri 4")

## SOP Median WT

tre.2.1 <- tre.2(Dates[1,], triage = "Tri P1")
tre.2.2 <- tre.2(Dates[1,], triage = "Tri P2")

## % SOP patient seen within WT

tre.2.1.1 <- tre.2.1(Dates[1,], triage = "Tri P1")
tre.2.1.2 <- tre.2.1(Dates[1,], triage = "Tri P2")

## MRSA Bacteremia in Acute Beds/1000 Acute Patient Days
tre.4 <- kpi.3.3(Dates[1,], trend = TRUE)



# Update reporting month and period -----------------------------------

addDataFrame(Series, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=2, startColumn=2)

# Write to Excel templates -----------------------------------------------

# 12 Months Trend

## % AE cases within pledged WT
addDataFrame(tre.1.1, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=2)
addDataFrame(tre.1.2, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=11, startColumn=2)
addDataFrame(tre.1.3, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=19, startColumn=2)
addDataFrame(tre.1.4, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=27, startColumn=2)

## SOP Median WT
for(i in 1:nrow(tre.2.1)){
    temp_df <- t(as.numeric(tre.2.1[i,]))
    temp_ri <- 35+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}


for(i in 1:nrow(tre.2.2)){
    temp_df <- t(as.numeric(tre.2.2[i,]))
    temp_ri <- 45+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}

## % SOP patient seen within WT
for(i in 1:nrow(tre.2.1.1)){
    temp_df <- t(as.numeric(tre.2.1.1[i,]))
    temp_ri <- 55+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}


for(i in 1:nrow(tre.2.1.2)){
    temp_df <- t(as.numeric(tre.2.1.2[i,]))
    temp_ri <- 65+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}

## MRSA Bacteremia in Acute Beds/1000 Acute Patient Days
for(i in 1:10){
    temp_df <- t(as.numeric(tre.4[i,]))
    temp_ri <- 91+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}

