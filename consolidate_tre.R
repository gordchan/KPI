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
tre.4.a <- kpi.3.3(Dates[1,], trend = TRUE)

## URR
tre.5.a <- tre.5(Dates[1,])

## % of stroke ever treated in ASU
tre.6.a <- tre.6(Dates[1,])

## % of hip fracture surgery done within 2 days
tre.7.a <- tre.7(Dates[1,])

## % of Day Surgery (DS) plus Same Day Surgery (SDS)
tre.9.a <- tre.9(Dates[1,])

## Bed
tre.10.a <- tre.10(Dates[1,], item = "Occ")
tre.10.b <- tre.10(Dates[1,], item = "ALOS")

## Radiology WT
tre.12.a <- tre.12(Dates[1,], Mode = "CT")
tre.12.b <- tre.12(Dates[1,], Mode = "MRI")
tre.12.c <- tre.12(Dates[1,], Mode = "US")
tre.12.d <- tre.12(Dates[1,], Mode = "MAMMO")

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
    temp_df <- t(as.numeric(tre.4.a[i,]))
    temp_ri <- 91+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}

## URR
temp_df <- t(as.numeric(tre.5.a[1,]))
temp_ri <- 102
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
for(i in 2:5){
    temp_df <- t(as.numeric(tre.5.a[i,]))
    temp_ri <- 103+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}
for(i in 6:8){
    temp_df <- t(as.numeric(tre.5.a[i,]))
    temp_ri <- 104+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}

## % of stroke ever treated in ASU
temp_df <- t(as.numeric(tre.6.a[1,]))
temp_ri <- 113
addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
for(i in 2:3){
    temp_df <- t(as.numeric(tre.6.a[i,]))
    temp_ri <- 114+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}
temp_df <- t(as.numeric(tre.6.a[4,]))
temp_ri <- 118
addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
for(i in 5:7){
    temp_df <- t(as.numeric(tre.6.a[i,]))
    temp_ri <- 116+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}


## % of hip fracture surgery done within 2 days
temp_df <- t(as.numeric(tre.7.a[1,]))
temp_ri <- 124
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
temp_df <- t(as.numeric(tre.7.a[2,]))
temp_ri <- 126
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
temp_df <- t(as.numeric(tre.7.a[3,]))
temp_ri <- 129
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
for(i in 4:6){
    temp_df <- t(as.numeric(tre.7.a[i,]))
    temp_ri <- 128+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}
    
## % of Day Surgery (DS) plus Same Day Surgery (SDS)
temp_df <- t(as.numeric(tre.9.a[1,]))
temp_ri <- 146
addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
for(i in 2:5){
    temp_df <- t(as.numeric(tre.9.a[i,]))
    temp_ri <- 147+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}
for(i in 6:8){
    temp_df <- t(as.numeric(tre.9.a[i,]))
    temp_ri <- 148+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}

## Bed
    ### Bed Occ
for(i in 1:nrow(tre.10.a)){
    temp_df <- t(as.numeric(tre.10.a[i,]))
    temp_ri <- 157+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}
    ### ALOS
temp_df <- t(as.numeric(tre.10.b[1,]))
temp_ri <- 168
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
for(i in 2:nrow(tre.10.b)){
    temp_df <- t(as.numeric(tre.10.b[i,]))
    temp_ri <- 169+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}
    
## Radiology WT
    ### CT
temp_df <- t(as.numeric(tre.12.a[1,]))
temp_ri <- 179
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
temp_df <- t(as.numeric(tre.12.a[2,]))
temp_ri <- 181
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
for(i in 3:4){
    temp_df <- t(as.numeric(tre.12.a[i,]))
    temp_ri <- 181+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}
for(i in 5:7){
    temp_df <- t(as.numeric(tre.12.a[i,]))
    temp_ri <- 182+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}

    ### MRI
temp_df <- t(as.numeric(tre.12.b[1,]))
temp_ri <- 190
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
temp_df <- t(as.numeric(tre.12.b[2,]))
temp_ri <- 192
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
temp_df <- t(as.numeric(tre.12.b[3,]))
temp_ri <- 195
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
for(i in 4:6){
    temp_df <- t(as.numeric(tre.12.b[i,]))
    temp_ri <- 194+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}
    
    ### US
temp_df <- t(as.numeric(tre.12.c[1,]))
temp_ri <- 201
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
temp_df <- t(as.numeric(tre.12.c[2,]))
temp_ri <- 203
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
for(i in 3:4){
    temp_df <- t(as.numeric(tre.12.c[i,]))
    temp_ri <- 203+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}
for(i in 5:7){
    temp_df <- t(as.numeric(tre.12.c[i,]))
    temp_ri <- 204+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}
    
    ### MAMMO
temp_df <- t(as.numeric(tre.12.d[1,]))
temp_ri <- 212
addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
temp_df <- t(as.numeric(tre.12.d[2,]))
temp_ri <- 214
addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
for(i in 3:4){
    temp_df <- t(as.numeric(tre.12.d[i,]))
    temp_ri <- 214+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}
for(i in 5:7){
    temp_df <- t(as.numeric(tre.12.d[i,]))
    temp_ri <- 215+i-1
    addDataFrame(temp_df, sheets.TRE$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
}