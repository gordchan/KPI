#
# Consolidate data for tre report
#

# Load and prepare TRE source data -----------------------------------

# 12 Months Trend

## % AE cases within pledged WT
tre.1.a <- tre.1(Dates[1,], triage = "Tri 1")
tre.1.b <- tre.1(Dates[1,], triage = "Tri 2")
tre.1.c <- tre.1(Dates[1,], triage = "Tri 3")
tre.1.d <- tre.1(Dates[1,], triage = "Tri 4")

## SOP Median WT

tre.2.a <- tre.2(Dates[1,], triage = "Tri P1")
tre.2.b <- tre.2(Dates[1,], triage = "Tri P2")

## % SOP patient seen within WT

tre.2.1.a <- tre.2.1(Dates[1,], triage = "Tri P1")
tre.2.1.b <- tre.2.1(Dates[1,], triage = "Tri P2")

## MRSA Bacteremia in Acute Beds/1000 Acute Patient Days
tre.4 <- kpi.3.3(Dates[1,], trend = TRUE)

## MRSA Bacteremia in Acute Beds/1000 Acute Patient Days
tre.5.a <- tre.5(Dates[1,])

# Update reporting month and period -----------------------------------

addDataFrame(Series, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=2, startColumn=2)

# Write to Excel templates -----------------------------------------------

# 12 Months Trend

## % AE cases within pledged WT
addDataFrame(tre.1.a, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=2)
addDataFrame(tre.1.b, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=11, startColumn=2)
addDataFrame(tre.1.c, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=19, startColumn=2)
addDataFrame(tre.1.d, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=27, startColumn=2)

## SOP Median WT
for(i in 1:nrow(tre.2.a)){
    temp_df <- t(as.numeric(tre.2.a[i,]))
    temp_ri <- 35+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}


for(i in 1:nrow(tre.2.b)){
    temp_df <- t(as.numeric(tre.2.b[i,]))
    temp_ri <- 45+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}

## % SOP patient seen within WT
for(i in 1:nrow(tre.2.1.a)){
    temp_df <- t(as.numeric(tre.2.1.a[i,]))
    temp_ri <- 55+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}


for(i in 1:nrow(tre.2.1.b)){
    temp_df <- t(as.numeric(tre.2.1.b[i,]))
    temp_ri <- 65+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}

## MRSA Bacteremia in Acute Beds/1000 Acute Patient Days
for(i in 1:10){
    temp_df <- t(as.numeric(tre.4[i,]))
    temp_ri <- 91+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}

## URR
temp_df <- t(as.numeric(tre.5[1,]))
temp_ri <- 102
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
temp_df <- t(as.numeric(tre.5[2:5,]))
temp_ri <- 104
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
temp_df <- t(as.numeric(tre.5[6:8,]))
temp_ri <- 109
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)

