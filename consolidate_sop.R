#
# Consolidate data for SOP report
#

# Load and prepare SOP source data -----------------------------------

# Current Year
    sop.ent.c <- kpi.2(Dates[1,], specialty = "ENT", index = 1:5)
    sop.gyn.c <- kpi.2(Dates[1,], specialty = "GYN", index = 1:5)
    sop.med.c <- kpi.2(Dates[1,], specialty = "MED", index = 1:5)
    sop.oph.c <- kpi.2(Dates[1,], specialty = "OPH", index = 1:5)
    sop.ort.c <- kpi.2(Dates[1,], specialty = "ORT", index = 1:5)
    sop.pae.c <- kpi.2(Dates[1,], specialty = "PAE", index = 1:5)
    sop.psy.c <- kpi.2(Dates[1,], specialty = "PSY", index = 1:5)
    sop.sur.c <- kpi.2(Dates[1,], specialty = "SUR", index = 1:5)
    sop.ovr.c <- kpi.2(Dates[1,], specialty = "Overall", index = 1:5)
# Previous Year	
    sop.ent.p <- kpi.2(Dates[3,], specialty = "ENT", index = 1:5)
    sop.gyn.p <- kpi.2(Dates[3,], specialty = "GYN", index = 1:5)
    sop.med.p <- kpi.2(Dates[3,], specialty = "MED", index = 1:5)
    sop.oph.p <- kpi.2(Dates[3,], specialty = "OPH", index = 1:5)
    sop.ort.p <- kpi.2(Dates[3,], specialty = "ORT", index = 1:5)
    sop.pae.p <- kpi.2(Dates[3,], specialty = "PAE", index = 1:5)
    sop.psy.p <- kpi.2(Dates[3,], specialty = "PSY", index = 1:5)
    sop.sur.p <- kpi.2(Dates[3,], specialty = "SUR", index = 1:5)
    sop.ovr.p <- kpi.2(Dates[3,], specialty = "Overall", index = 1:5)
    
# Update reporting month and period -----------------------------------

# addDataFrame(dates, sheets.KPI$interface, col.names=FALSE, row.names=FALSE, startRow=1, startColumn=2)

# Write to Excel templates -----------------------------------------------
    
# Targets
    
#     for (i in 1:25){
#         kpi.t.i <- kpi.t(i, Dates[1,])
#         addDataFrame(kpi.t.i, sheets.KPI$interface, col.names=FALSE, row.names=FALSE, startRow=(3+i), startColumn=13)
#     }

# Current Year

    addDataFrame(sop.ent.c[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=2)
    addDataFrame(sop.gyn.c[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=5, startColumn=2)
    addDataFrame(sop.med.c[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=7, startColumn=2)
    addDataFrame(sop.oph.c[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=9, startColumn=2)
    addDataFrame(sop.ort.c[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=11, startColumn=2)
    addDataFrame(sop.pae.c[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=13, startColumn=2)
    addDataFrame(sop.psy.c[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=15, startColumn=2)
    addDataFrame(sop.sur.c[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=17, startColumn=2)
    addDataFrame(sop.ovr.c[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=19, startColumn=2)
    
    addDataFrame(sop.ent.c[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=23, startColumn=2)
    addDataFrame(sop.gyn.c[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=26, startColumn=2)
    addDataFrame(sop.med.c[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=29, startColumn=2)
    addDataFrame(sop.oph.c[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=32, startColumn=2)
    addDataFrame(sop.ort.c[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=35, startColumn=2)
    addDataFrame(sop.pae.c[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=38, startColumn=2)
    addDataFrame(sop.psy.c[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=41, startColumn=2)
    addDataFrame(sop.sur.c[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=44, startColumn=2)
    addDataFrame(sop.ovr.c[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=47, startColumn=2)

# Previous Year

    addDataFrame(sop.ent.p[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=3, startColumn=12)
    addDataFrame(sop.gyn.p[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=5, startColumn=12)
    addDataFrame(sop.med.p[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=7, startColumn=12)
    addDataFrame(sop.oph.p[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=9, startColumn=12)
    addDataFrame(sop.ort.p[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=11, startColumn=12)
    addDataFrame(sop.pae.p[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=13, startColumn=12)
    addDataFrame(sop.psy.p[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=15, startColumn=12)
    addDataFrame(sop.sur.p[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=17, startColumn=12)
    addDataFrame(sop.ovr.p[c(3:4),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=19, startColumn=12)
    
    addDataFrame(sop.ent.p[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=23, startColumn=12)
    addDataFrame(sop.gyn.p[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=26, startColumn=12)
    addDataFrame(sop.med.p[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=29, startColumn=12)
    addDataFrame(sop.oph.p[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=32, startColumn=12)
    addDataFrame(sop.ort.p[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=35, startColumn=12)
    addDataFrame(sop.pae.p[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=38, startColumn=12)
    addDataFrame(sop.psy.p[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=41, startColumn=12)
    addDataFrame(sop.sur.p[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=44, startColumn=12)
    addDataFrame(sop.ovr.p[c(1:2,5),], sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=47, startColumn=12)

