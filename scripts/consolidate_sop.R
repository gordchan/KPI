#
# Consolidate data for SOP report
#

# Load and prepare SOP source data -----------------------------------

# Current Year
    sop.ent.c <- kpi.2(Dates[1,], specialty = "ENT", index = 1:5)
    sop.gyn.c <- kpi.2(Dates[1,], specialty = "GYN", index = 1:5)
    sop.med.c <- kpi.2(Dates[1,], specialty = "MED", index = 1:5)
    sop.oph.c <- kpi.2(Dates[1,], specialty = "OPH", index = 1:5)
        sop.oph.c.p1 <- kpi.2(Dates[1,], specialty = "OPH", index = 3)
            for (i in 1:ncol(sop.oph.c.p1)){if(sop.oph.c.p1[,i]!="N.A." & is.character(sop.oph.c.p1[,i])){sop.oph.c.p1[,i] <- as.numeric(sop.oph.c.p1[,i])}}
        sop.oph.c.p2 <- kpi.2(Dates[1,], specialty = "OPH", index = 4)
            for (i in 1:ncol(sop.oph.c.p2)){if(sop.oph.c.p2[,i]!="N.A." & is.character(sop.oph.c.p2[,i])){sop.oph.c.p2[,i] <- as.numeric(sop.oph.c.p2[,i])}}
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
        sop.oph.p.p1 <- kpi.2(Dates[3,], specialty = "OPH", index = 3)
            for (i in 1:ncol(sop.oph.p.p1)){if(sop.oph.p.p1[,i]!="N.A." & is.character(sop.oph.p.p1[,i])){sop.oph.p.p1[,i] <- as.numeric(sop.oph.p.p1[,i])}}
        sop.oph.p.p2 <- kpi.2(Dates[3,], specialty = "OPH", index = 4)
            for (i in 1:ncol(sop.oph.p.p2)){if(sop.oph.p.p2[,i]!="N.A." & is.character(sop.oph.p.p2[,i])){sop.oph.p.p2[,i] <- as.numeric(sop.oph.p.p2[,i])}}
    sop.ort.p <- kpi.2(Dates[3,], specialty = "ORT", index = 1:5)
    sop.pae.p <- kpi.2(Dates[3,], specialty = "PAE", index = 1:5)
    sop.psy.p <- kpi.2(Dates[3,], specialty = "PSY", index = 1:5)
    sop.sur.p <- kpi.2(Dates[3,], specialty = "SUR", index = 1:5)
    sop.ovr.p <- kpi.2(Dates[3,], specialty = "Overall", index = 1:5)
    
# Update reporting month and period -----------------------------------

addDataFrame(Series, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=53, startColumn=2)
    
# Get SOP WT Trend by specialty  -------------------------------------
    
    sop.ent.t <- kpi.11(Dates[1,], specialty = "ENT")
    sop.gyn.t <- kpi.11(Dates[1,], specialty = "GYN")
    sop.med.t <- kpi.11(Dates[1,], specialty = "MED")
    sop.oph.t <- kpi.11(Dates[1,], specialty = "OPH")
    sop.ort.t <- kpi.11(Dates[1,], specialty = "ORT")
    sop.pae.t <- kpi.11(Dates[1,], specialty = "PAE")
    sop.psy.t <- kpi.11(Dates[1,], specialty = "PSY")
    sop.sur.t <- kpi.11(Dates[1,], specialty = "SUR")
    
    sop.seu.t <- kpi.11.2(Dates[1,])
    
    sop.uro.t <- kpi.11.3(Dates[1,])
    
# Index table for filling SOP trend data ---------------------------------
    
    sop.t.index <- data.frame(
        Inst = c("CMC", "KCH", "KWH", "NLTH", "OLM", "PMH", "YCH", "KWC", "HA"),
        ENT = c(54:62),
        GYN = c(64:72),
        MED = c(74:82),
        OPH = c(84:92),
        ORT = c(94:102),
        PAE = c(104:112),
        PSY = c(114:122),
        SUR = c(124:132),
        SURexUROL = c(134:142),
        UROL = c(144:152),
        stringsAsFactors = FALSE
    )
    
    uro.t.index <- data.frame(
        Cluster = c("HKEC", "HKWC", "KCC", "KEC", "KWC", "NTEC", "NTWC", "HA"),
        UROL.c = c(154:161),
        stringsAsFactors = FALSE
    )
    
    
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
        addDataFrame(sop.oph.c.p1, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=9, startColumn=2)
        addDataFrame(sop.oph.c.p2, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=10, startColumn=2)
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
        addDataFrame(sop.oph.p.p1, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=9, startColumn=12)
        addDataFrame(sop.oph.p.p2, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=10, startColumn=12)
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

# SOP WT Trend
    
    # ENT
    for(i in 1:nrow(sop.t.index)){
        if(sop.t.index$Inst[i] %in% row.names(sop.ent.t)){
                        temp_df <- t(as.numeric(sop.ent.t[which(row.names(sop.ent.t)==sop.t.index$Inst[i]),]))
            temp_ri <- sop.t.index %>% select(Inst, ENT) %>% filter(Inst == sop.t.index$Inst[i])
                temp_ri <- temp_ri[1,2]
        addDataFrame(temp_df, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
        }
    }

    # GYN
    for(i in 1:nrow(sop.t.index)){
        if(sop.t.index$Inst[i] %in% row.names(sop.gyn.t)){
                        temp_df <- t(as.numeric(sop.gyn.t[which(row.names(sop.gyn.t)==sop.t.index$Inst[i]),]))
        temp_ri <- sop.t.index %>% select(Inst, GYN) %>% filter(Inst == sop.t.index$Inst[i])
        temp_ri <- temp_ri[1,2]
        addDataFrame(temp_df, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
        }
    }
    
    # MED
    for(i in 1:nrow(sop.t.index)){
        if(sop.t.index$Inst[i] %in% row.names(sop.med.t)){
                        temp_df <- t(as.numeric(sop.med.t[which(row.names(sop.med.t)==sop.t.index$Inst[i]),]))
        temp_ri <- sop.t.index %>% select(Inst, MED) %>% filter(Inst == sop.t.index$Inst[i])
        temp_ri <- temp_ri[1,2]
        addDataFrame(temp_df, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
        }
    }
    
    # OPH
    for(i in 1:nrow(sop.t.index)){
        if(sop.t.index$Inst[i] %in% row.names(sop.oph.t)){
                    temp_df <- t(as.numeric(sop.oph.t[which(row.names(sop.oph.t)==sop.t.index$Inst[i]),]))
            temp_ri <- sop.t.index %>% select(Inst, OPH) %>% filter(Inst == sop.t.index$Inst[i])
            temp_ri <- temp_ri[1,2]
            addDataFrame(temp_df, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
        }
    }
    
    # ORT
    for(i in 1:nrow(sop.t.index)){
        if(sop.t.index$Inst[i] %in% row.names(sop.ort.t)){
                    temp_df <- t(as.numeric(sop.ort.t[which(row.names(sop.ort.t)==sop.t.index$Inst[i]),]))
            temp_ri <- sop.t.index %>% select(Inst, ORT) %>% filter(Inst == sop.t.index$Inst[i])
            temp_ri <- temp_ri[1,2]
            addDataFrame(temp_df, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
        }
    }
    
    # PAE
    for(i in 1:nrow(sop.t.index)){
        if(sop.t.index$Inst[i] %in% row.names(sop.pae.t)){
                    temp_df <- t(as.numeric(sop.pae.t[which(row.names(sop.pae.t)==sop.t.index$Inst[i]),]))
            temp_ri <- sop.t.index %>% select(Inst, PAE) %>% filter(Inst == sop.t.index$Inst[i])
            temp_ri <- temp_ri[1,2]
            addDataFrame(temp_df, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
        }
    }
    
    # PSY
    for(i in 1:nrow(sop.t.index)){
        if(sop.t.index$Inst[i] %in% row.names(sop.psy.t)){
                    temp_df <- t(as.numeric(sop.psy.t[which(row.names(sop.psy.t)==sop.t.index$Inst[i]),]))
            temp_ri <- sop.t.index %>% select(Inst, PSY) %>% filter(Inst == sop.t.index$Inst[i])
            temp_ri <- temp_ri[1,2]
            addDataFrame(temp_df, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
        }
    }
    
    # SUR
    for(i in 1:nrow(sop.t.index)){
        if(sop.t.index$Inst[i] %in% row.names(sop.sur.t)){
                    temp_df <- t(as.numeric(sop.sur.t[which(row.names(sop.sur.t)==sop.t.index$Inst[i]),]))
            temp_ri <- sop.t.index %>% select(Inst, SUR) %>% filter(Inst == sop.t.index$Inst[i])
            temp_ri <- temp_ri[1,2]
            addDataFrame(temp_df, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
        }
    }
    
    # SUR excl UROL
    for(i in 1:nrow(sop.t.index)){
        if(sop.t.index$Inst[i] %in% row.names(sop.seu.t)){
            temp_df <- sop.seu.t[which(row.names(sop.seu.t)==sop.t.index$Inst[i]),]
            temp_ri <- sop.t.index %>% select(Inst, SURexUROL) %>% filter(Inst == sop.t.index$Inst[i])
            temp_ri <- temp_ri[1,2]
            addDataFrame(temp_df, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
        }
    }
    
    # UROL
    for(i in 1:nrow(sop.t.index)){
        if(sop.t.index$Inst[i] %in% row.names(sop.uro.t)){
            temp_df <- sop.uro.t[which(row.names(sop.uro.t)==sop.t.index$Inst[i]),]
            temp_ri <- sop.t.index %>% select(Inst, UROL) %>% filter(Inst == sop.t.index$Inst[i])
            temp_ri <- temp_ri[1,2]
            addDataFrame(temp_df, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
        }
    }
    
    # UROL by Clusters
    for(i in 1:nrow(uro.t.index)){
        if(uro.t.index$Cluster[i] %in% row.names(sop.uro.t)){
            temp_df <- sop.uro.t[which(row.names(sop.uro.t)==uro.t.index$Cluster[i]),]
            temp_ri <- uro.t.index %>% select(Cluster, UROL.c) %>% filter(Cluster == uro.t.index$Cluster[i])
            temp_ri <- temp_ri[1,2]
            addDataFrame(temp_df, sheets.SOP$interface, col.names=FALSE, row.names=FALSE, startRow=temp_ri, startColumn=2)
        }
    }