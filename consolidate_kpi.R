#
# Consolidate data for KPI report
#

# Load and prepare KPI source data -----------------------------------

# Current Year

    kpi.1.c <- kpi.1(to.MmmYY)
    kpi.2.c <- kpi.2(to.MmmYY)
    kpi.3.1.c <- kpi.3("kpi.3.1", to.MmmYY)
    kpi.3.2.c <- kpi.3("kpi.3.2", to.MmmYY)	
    kpi.3.3.c <- kpi.3.3(to.MmmYY)	
    kpi.3.4.c <- kpi.10(to.MmmYY)	
    kpi.4.1.c <- kpi.4.1(to.MmmYY)	
    kpi.4.2.c <- kpi.4.2(to.MmmYY)	
    kpi.4.3.c <- kpi.4.3(to.MmmYY)	
    kpi.5.1.c <- kpi.5.1(to.MmmYY)	
    kpi.5.2.c <- kpi.5("kpi.5.2", to.MmmYY)	
    kpi.5.3.c <- kpi.5("kpi.5.3", to.MmmYY)	
    kpi.6.1.c <- kpi.6("kpi.6.1", to.MmmYY)	
    kpi.6.2.c <- kpi.6("kpi.6.2", to.MmmYY)
    kpi.7.1.c <- kpi.7("kpi.7.1", to.MmmYY)	
    kpi.7.2.c <- kpi.7("kpi.7.2", to.MmmYY)	
    kpi.7.3.c <- kpi.7("kpi.7.3", to.MmmYY)	
    kpi.7.4.c <- kpi.7("kpi.7.4", to.MmmYY)	

# Previous Year	

    kpi.1.p <- kpi.1(prev.to.MmmYY)
    kpi.2.p <- kpi.2(prev.to.MmmYY)
    kpi.3.1.p <- kpi.3("kpi.3.1", prev.to.MmmYY)
    kpi.3.2.p <- kpi.3("kpi.3.2", prev.to.MmmYY)
    kpi.3.3.p <- kpi.3.3(prev.to.MmmYY)
    kpi.3.4.p <- kpi.10(prev.to.MmmYY)
    kpi.4.1.p <- kpi.4.1(prev.to.MmmYY)
    kpi.4.2.p <- kpi.4.2(prev.to.MmmYY)
    kpi.4.3.p <- kpi.4.3(prev.to.MmmYY)
    kpi.5.1.p <- kpi.5.1(prev.to.MmmYY)
    kpi.5.2.p <- kpi.5("kpi.5.2", prev.to.MmmYY)
    kpi.5.3.p <- kpi.5("kpi.5.3", prev.to.MmmYY)
    kpi.6.1.p <- kpi.6("kpi.6.1", prev.to.MmmYY)
    kpi.6.2.p <- kpi.6("kpi.6.2", prev.to.MmmYY)
    kpi.7.1.p <- kpi.7("kpi.7.1", prev.to.MmmYY)
    kpi.7.2.p <- kpi.7("kpi.7.2", prev.to.MmmYY)
    kpi.7.3.p <- kpi.7("kpi.7.3", prev.to.MmmYY)
    kpi.7.4.p <- kpi.7("kpi.7.4", prev.to.MmmYY)

# Update reporting month and period -----------------------------------

addDataFrame(dates, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=1, startColumn=2)

# Write to Excel templates -----------------------------------------------
    
# Targets
    
    for (i in 1:25){
        kpi.t.i <- kpi.t(i, to.MmmYY)
        addDataFrame(kpi.t.i, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=(3+i), startColumn=13)
    }
    
# Current Year
    
    addDataFrame(kpi.1.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=4, startColumn=3)
    addDataFrame(kpi.2.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=8, startColumn=3)
    addDataFrame(kpi.3.1.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=13, startColumn=3)
    addDataFrame(kpi.3.2.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=14, startColumn=3)
    addDataFrame(kpi.3.3.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=15, startColumn=3)
    addDataFrame(kpi.3.4.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=16, startColumn=3)
    addDataFrame(kpi.4.1.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=17, startColumn=3)
    addDataFrame(kpi.4.2.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=18, startColumn=3)
    addDataFrame(kpi.4.3.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=19, startColumn=3)
    addDataFrame(kpi.5.1.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=20, startColumn=3)
    addDataFrame(kpi.5.2.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=21, startColumn=3)
    addDataFrame(kpi.5.3.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=22, startColumn=3)
    addDataFrame(kpi.6.1.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=23, startColumn=3)
    addDataFrame(kpi.6.2.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=24, startColumn=3)
    addDataFrame(kpi.7.1.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=25, startColumn=3)
    addDataFrame(kpi.7.2.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=26, startColumn=3)
    addDataFrame(kpi.7.3.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=27, startColumn=3)
    addDataFrame(kpi.7.4.c, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=28, startColumn=3)
    
# Previous Year
    
    addDataFrame(kpi.1.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=4, startColumn=13)
    addDataFrame(kpi.2.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=8, startColumn=13)
    addDataFrame(kpi.3.1.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=13, startColumn=13)
    addDataFrame(kpi.3.2.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=14, startColumn=13)
    addDataFrame(kpi.3.3.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=15, startColumn=13)
    addDataFrame(kpi.3.4.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=16, startColumn=13)
    addDataFrame(kpi.4.1.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=17, startColumn=13)
    addDataFrame(kpi.4.2.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=18, startColumn=13)
    addDataFrame(kpi.4.3.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=19, startColumn=13)
    addDataFrame(kpi.5.1.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=20, startColumn=13)
    addDataFrame(kpi.5.2.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=21, startColumn=13)
    addDataFrame(kpi.5.3.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=22, startColumn=13)
    addDataFrame(kpi.6.1.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=23, startColumn=13)
    addDataFrame(kpi.6.2.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=24, startColumn=13)
    addDataFrame(kpi.7.1.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=25, startColumn=13)
    addDataFrame(kpi.7.2.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=26, startColumn=13)
    addDataFrame(kpi.7.3.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=27, startColumn=13)
    addDataFrame(kpi.7.4.p, sheets.KPI$source, col.names=FALSE, row.names=FALSE, startRow=28, startColumn=13)
    