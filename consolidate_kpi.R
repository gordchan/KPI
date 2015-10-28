#
# Consolidate data for KPI report
#

# Load and prepare KPI NEW_INTERFACE data -----------------------------------

# Current Year

    kpi.1.c <- kpi.1(Dates[1,])
    kpi.2.c <- kpi.2(Dates[1,])
    kpi.3.1.c <- kpi.3("kpi.3.1", Dates[1,])
    kpi.3.2.c <- kpi.3("kpi.3.2", Dates[1,])
    kpi.3.3.c <- kpi.3.3(Dates[1,])
    kpi.3.4.c <- kpi.10(Dates[1,])	
    kpi.4.1.c <- kpi.4.1(Dates[1,])	
    kpi.4.2.c <- kpi.4.2(Dates[1,])	
    kpi.4.3.c <- kpi.4.3(Dates[1,])	
    kpi.5.1.c <- kpi.5.1(Dates[1,])	
    kpi.5.2.c <- kpi.5("kpi.5.2", Dates[1,])	
    kpi.5.3.c <- kpi.5("kpi.5.3", Dates[1,])	
    kpi.6.1.c <- kpi.6("kpi.6.1", Dates[1,])	
    kpi.6.2.c <- kpi.6("kpi.6.2", Dates[1,])
    kpi.7.1.c <- kpi.7("kpi.7.1", Dates[1,])	
    kpi.7.2.c <- kpi.7("kpi.7.2", Dates[1,])	
    kpi.7.3.c <- kpi.7("kpi.7.3", Dates[1,])	
    kpi.7.4.c <- kpi.7("kpi.7.4", Dates[1,])	

# Previous Year	

    kpi.1.p <- kpi.1(Dates[3,])
    kpi.2.p <- kpi.2(Dates[3,])
    kpi.3.1.p <- kpi.3("kpi.3.1", Dates[3,])
    kpi.3.2.p <- kpi.3("kpi.3.2", Dates[3,])
    kpi.3.3.p <- kpi.3.3(Dates[3,])
    kpi.3.4.p <- kpi.10(Dates[3,])
    kpi.4.1.p <- kpi.4.1(Dates[3,])
    kpi.4.2.p <- kpi.4.2(Dates[3,])
    kpi.4.3.p <- kpi.4.3(Dates[3,])
    kpi.5.1.p <- kpi.5.1(Dates[3,])
    kpi.5.2.p <- kpi.5("kpi.5.2", Dates[3,])
    kpi.5.3.p <- kpi.5("kpi.5.3", Dates[3,])
    kpi.6.1.p <- kpi.6("kpi.6.1", Dates[3,])
    kpi.6.2.p <- kpi.6("kpi.6.2", Dates[3,])
    kpi.7.1.p <- kpi.7("kpi.7.1", Dates[3,])
    kpi.7.2.p <- kpi.7("kpi.7.2", Dates[3,])
    kpi.7.3.p <- kpi.7("kpi.7.3", Dates[3,])
    kpi.7.4.p <- kpi.7("kpi.7.4", Dates[3,])
    
# DS + SDS
    
    kpi.8.ovr <- kpi.8(Dates[1,], specialty = "Overall", inst = "KWC")
    kpi.8.ovr.HA <- kpi.8(Dates[1,], specialty = "Overall", inst = "HA")
    kpi.8.ent <- kpi.8(Dates[1,], specialty = "ENT", inst = "KWC")
    kpi.8.ent.HA <- kpi.8(Dates[1,], specialty = "ENT", inst = "HA")
    kpi.8.og <- kpi.8(Dates[1,], specialty = "OG", inst = "KWC")
    kpi.8.og.HA <- kpi.8(Dates[1,], specialty = "OG", inst = "HA")
    kpi.8.oph <- kpi.8(Dates[1,], specialty = "OPH", inst = "KWC")
    kpi.8.oph.HA <- kpi.8(Dates[1,], specialty = "OPH", inst = "HA")
    kpi.8.ort <- kpi.8(Dates[1,], specialty = "ORT", inst = "KWC")
    kpi.8.ort.HA <- kpi.8(Dates[1,], specialty = "ORT", inst = "HA")
    kpi.8.sur <- kpi.8(Dates[1,], specialty = "SUR", inst = "KWC")
    kpi.8.sur.HA <- kpi.8(Dates[1,], specialty = "SUR", inst = "HA")

# Update reporting month and period -----------------------------------

#addDataFrame(dates, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=1, startColumn=2)

# Write to Excel templates -----------------------------------------------
    
# Targets
    
#     for (i in 1:25){
#         kpi.t.i <- kpi.t(i, Dates[1,])
#         addDataFrame(kpi.t.i, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=(3+i), startColumn=13)
#     }

# Current Year
    
    addDataFrame(kpi.1.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=3)
    addDataFrame(kpi.2.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=7, startColumn=3)
    addDataFrame(kpi.3.1.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=12, startColumn=3)
    addDataFrame(kpi.3.2.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=13, startColumn=3)
    addDataFrame(kpi.3.3.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=14, startColumn=3)
    addDataFrame(kpi.3.4.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=18, startColumn=3)
    addDataFrame(kpi.4.1.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=21, startColumn=3)
    addDataFrame(kpi.4.2.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=24, startColumn=3)
    addDataFrame(kpi.4.3.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=25, startColumn=3)
    addDataFrame(kpi.5.1.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=28, startColumn=3)
    addDataFrame(kpi.5.2.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=29, startColumn=3)
    addDataFrame(kpi.5.3.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=30, startColumn=3)
    addDataFrame(kpi.6.1.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=26, startColumn=3)
    addDataFrame(kpi.6.2.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=27, startColumn=3)
    addDataFrame(kpi.7.1.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=31, startColumn=3)
    addDataFrame(kpi.7.2.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=32, startColumn=3)
    addDataFrame(kpi.7.3.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=33, startColumn=3)
    addDataFrame(kpi.7.4.c, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=34, startColumn=3)
    
# Previous Year
    
    addDataFrame(kpi.1.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=13)
    addDataFrame(kpi.2.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=7, startColumn=13)
    addDataFrame(kpi.3.1.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=12, startColumn=13)
    addDataFrame(kpi.3.2.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=13, startColumn=13)
    addDataFrame(kpi.3.3.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=14, startColumn=13)
    addDataFrame(kpi.3.4.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=18, startColumn=13)
    addDataFrame(kpi.4.1.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=21, startColumn=13)
    addDataFrame(kpi.4.2.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=24, startColumn=13)
    addDataFrame(kpi.4.3.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=25, startColumn=13)
    addDataFrame(kpi.5.1.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=28, startColumn=13)
    addDataFrame(kpi.5.2.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=29, startColumn=13)
    addDataFrame(kpi.5.3.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=30, startColumn=13)
    addDataFrame(kpi.6.1.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=26, startColumn=13)
    addDataFrame(kpi.6.2.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=27, startColumn=13)
    addDataFrame(kpi.7.1.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=31, startColumn=13)
    addDataFrame(kpi.7.2.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=32, startColumn=13)
    addDataFrame(kpi.7.3.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=33, startColumn=13)
    addDataFrame(kpi.7.4.p, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=34, startColumn=13)

# DS + SDS
    
    addDataFrame(kpi.8.ovr, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=2, startColumn=24)
    addDataFrame(kpi.8.ovr.HA, sheets.KPI$NEW_INTERFACE, col.names=FALSE, row.names=FALSE, startRow=2, startColumn=34)
    