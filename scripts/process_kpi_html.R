# Process KPI-HTML
#
# Direct from HTML version of functions.
#
# Gordon CHAN
#
# Mar 2016

# 15 items that are natively in HTML
#kpi.3.1 & 2
    # kpi.3.1.1
    # kpi.3.1.2
    # kpi.3.2.1
    # kpi.3.2.2
#kpi.3.3
    # kpi.3.3
#kpi.4
    # kpi.4.1
    # kpi.4.2
#kpi.8
    # kpi.8
    # kpi.5.1
#tre.3.1 & 2
    # tre.3.1.1
    # tre.3.1.2
    # tre.3.2.1
    # tre.3.2.2
#tre.6
    # tre.6
#tre.7
    # tre.7
#tre.9
    # tre.9

# kpi.3 ---------------------------------------------------

kpi.3.html <- function(kpi, Mmm){
    
    # Announce fx started
    
    message("[kpi.3] Function started ", Sys.time())
    
    kpi_source_helper(Mmm)
    
    SAR.frame <- empty.frame
    SAR.frame[1,] <- NA
    
    if(kpi=="kpi.3.1"){
        path.1st <- grep("(.*kpi.3.1.1 .*html)", source.kpi$filepath, value = TRUE)
        path.adm <- grep("(.*kpi.3.1.2 .*html)", source.kpi$filepath, value = TRUE)
        
        SAR.1st <- read_Chtml(path.1st, 1:2)
        SAR.adm <- read_Chtml(path.adm, 1:2)
        
    }else if(kpi=="kpi.3.2"){
        path.1st <- grep("(.*kpi.3.2.1 .*)", source.kpi$filepath, value = TRUE)
        path.adm <- grep("(.*kpi.3.2.2 .*)", source.kpi$filepath, value = TRUE)
        
        SAR.1st <- read_Chtml(path.1st, 1:2)
        SAR.adm <- read_Chtml(path.adm, 1:2)
        
    }
    
    names(SAR.1st) <- gsub("(Row Total)(.*)", "HA\\2", names(SAR.1st))
    names(SAR.adm) <- gsub("(Row Total)(.*)", "HA\\2", names(SAR.adm))
    
    names(SAR.1st) <- gsub("^(Institution .* >> )", "", names(SAR.1st))
    names(SAR.adm) <- gsub("^(Institution .* >> )", "", names(SAR.adm))

    names(SAR.1st) <- gsub("( >> No. of A&E 1st Attendances)$", "", names(SAR.1st))
    names(SAR.adm) <- gsub("( >> No. of A&E 1st Attendances)$", "", names(SAR.adm))
        
    names(SAR.1st) <- gsub("( \\(.*\\))", "", names(SAR.1st))
    names(SAR.adm) <- gsub("( \\(.*\\))", "", names(SAR.adm))
    
    names(SAR.1st) <- gsub("( $)", "", names(SAR.1st))
    names(SAR.adm) <- gsub("( $)", "", names(SAR.adm))
    
    names(SAR.1st) <- gsub("( )", "_", names(SAR.1st))
    names(SAR.adm) <- gsub("( )", "_", names(SAR.adm))
    
    for (i in 1:(ncol(SAR.1st)-4)){
        SAR.1st[,i+4] <- as.numeric(SAR.1st[,i+4])
        SAR.adm[,i+4] <- as.numeric(SAR.adm[,i+4])
    }
    
    # Substitute NA with 0
    
    SAR.1st[is.na(SAR.1st)] <- 0 
    SAR.adm[is.na(SAR.adm)] <- 0
    
    # Select KWC hospitals & HA
    
    SAR.1st <- SAR.1st %>% filter(!grepl("Total", Admission_Age)) %>%
        mutate(KWC = CMC + KWH + NLTH + PMH + YCH) %>%
        select(Admission_Age, Sex, Ambulance_Case, Triage_Category, CMC.1 = CMC, KWH.1 = KWH, NLTH.1 = NLTH, PMH.1 = PMH, YCH.1 = YCH, KWC.1 = KWC, HA.1 = HA)
    
    SAR.adm <- SAR.adm %>% filter(!grepl("Total", Admission_Age)) %>%
        mutate(KWC = CMC + KWH + NLTH + PMH + YCH) %>%
        select(Admission_Age, Sex, Ambulance_Case, Triage_Category, CMC.a = CMC, KWH.a = KWH, NLTH.a = NLTH, PMH.a = PMH, YCH.a = YCH, KWC.a = KWC, HA.a = HA)
    
    SAR <- suppressMessages(full_join(SAR.1st, SAR.adm))
    
    for (i in 1:(ncol(SAR)-4)){
        SAR[,i+4] <- sapply(SAR[,i+4], FUN = function(x){ifelse(is.na(x), 0, x)})
    }
    
    SAR.req <- SAR %>% select(-Admission_Age, -Sex, -Ambulance_Case, -Triage_Category) %>% mutate(HA = HA.a/HA.1) %>%
        mutate(CMC = CMC.1 * HA, KWH = KWH.1 * HA,  NLTH = NLTH.1 * HA, PMH = PMH.1 * HA, YCH = YCH.1 * HA, KWC = KWC.1 * HA) %>%
        summarise_each(funs(sum))
    
    SAR.frame$HA <- SAR.req$HA.a/SAR.req$HA.1
    SAR.frame$CMC <- SAR.req$CMC.a/SAR.req$CMC*SAR.frame$HA
    SAR.frame$KWH <- SAR.req$KWH.a/SAR.req$KWH*SAR.frame$HA
    SAR.frame$NLTH <- SAR.req$NLTH.a/SAR.req$NLTH*SAR.frame$HA
    SAR.frame$PMH <- SAR.req$PMH.a/SAR.req$PMH*SAR.frame$HA
    SAR.frame$YCH <- SAR.req$YCH.a/SAR.req$YCH*SAR.frame$HA
    SAR.frame$KWC <- SAR.req$KWC.a/SAR.req$KWC*SAR.frame$HA
    
    # Replace NA with N.A. for use in Excel
    
    SAR.prod <- data.frame(lapply(SAR.frame, FUN = function(x){ifelse(is.na(x), "N.A.", x)}), stringsAsFactors = FALSE)
    
    for (i in 1:ncol(SAR.prod)){
        if ("N.A." %in% SAR.prod[,i]){
        } else {
            SAR.prod[,i] <- as.numeric(SAR.prod[,i])
        }
    }
    
    SAR.prod
    
}




# kpi.4 --------------------------------------------------------


kpi.4.1.html <- function(Mmm){
    
    # Announce fx started
    
    message("[kpi.4.1] Function started ", Sys.time())
    
    kpi_source_helper(Mmm)
    
    path <- grep("(.*kpi.4.1 .*html)", source.kpi$filepath, value = TRUE)
    
    Stroke.frame <- empty.frame
    Stroke.frame[1,] <- NA
    
    
    Stroke <- read_Chtml(path, 1:2)
    
    names(Stroke) <- c("cluster", "institution", "ASU", "other", "total")
    
    Stroke$institution <- gsub("^Subtotal.*", "Subtotal", Stroke$institution)
    Stroke$cluster <- gsub("^Grand Total.*", "HA", Stroke$cluster)
    Stroke$institution <- gsub("^Grand Total.*", "HA", Stroke$institution)
    
    Stroke <- Stroke %>% filter(cluster == "KW" | cluster == "HA" | institution == "Subtotal")

    Stroke$institution <- gsub("Subtotal", NA, Stroke$institution)
    
    Stroke$cluster <- gsub(" $", "", Stroke$cluster)
    Stroke$cluster <- sapply(Stroke$cluster, FUN = function(x){ifelse(x=="HA", x, paste(x,"C", sep = ""))})
    
    for (i in 1:nrow(Stroke)){
        if (is.na(Stroke$institution[i])){
            Stroke$institution[i] <- Stroke$cluster[i]
        }
        if (is.na(Stroke$ASU[i]) & Stroke$other[i]!=0){
            Stroke$ASU[i] <- 0
        }
    }
    
    for (i in 1:3){
        Stroke[,i+2] <- as.numeric(Stroke[,i+2])
    }
    
    Stroke <- Stroke %>% mutate(ASU.rate = ASU/total) %>%
        select(institution, ASU.rate) %>%
        t() %>% as.data.frame(stringsAsFactors = FALSE)
    
    names(Stroke) <- Stroke[1,]
    
    Stroke <- Stroke[-1,]
    
    Stroke <- data.frame(lapply(Stroke[1,], FUN = function(x) as.numeric(x)))
    
    
    for (i in 1:length(Stroke.frame)){
        if (names(Stroke.frame)[i] %in% names(Stroke)){
            Stroke.frame[i] <- Stroke[names(Stroke.frame)[i]]
        }
    }
    
    # Replace NA with N.A. for use in Excel
    
#     Stroke.prod <- lapply(Stroke.frame, function(x){ifelse(is.na(x), "N.A.", x)})
#     Stroke.prod <- data.frame(Stroke.prod)
    
    # Return production dataframe
    
    Stroke.prod <- Stroke.frame
    
    Stroke.prod
    
}

kpi.4.2.html <- function(Mmm){
    
    # Announce fx started
    
    message("[kpi.4.2] Function started ", Sys.time())
    
    kpi_source_helper(Mmm)
    
    path <- grep("(.*kpi.4.2 .*html)", source.kpi$filepath, value = TRUE)
    
    Hip.frame <- empty.frame
    Hip.frame[1,] <- NA
    
    
    Hip <- read_Chtml(path, 1:2)
    
    names(Hip) <- c("cluster", "institution", "within_2", "between_2_to_4", "other", "total")
    
    Hip$institution <- gsub("^Subtotal.*", "Subtotal", Hip$institution)
    Hip$cluster <- gsub("^Grand.*", "HA", Hip$cluster)
    Hip$institution <- gsub("^Grand.*", "HA", Hip$institution)
    
    Hip <- Hip%>% filter(cluster == "KW" | cluster == "HA" | institution == "Subtotal")
    
    Hip$institution <- gsub(".*total.*", NA, Hip$institution)
    
    Hip$cluster <- gsub(" $", "", Hip$cluster)
    Hip$cluster <- sapply(Hip$cluster, FUN = function(x){ifelse(x=="HA", x, paste(x,"C", sep = ""))})
    
    for (i in 1:nrow(Hip)){
        if (is.na(Hip$institution[i])){
            Hip$institution[i] <- Hip$cluster[i]
        }
        if (is.na(Hip$within_2[i]) & Hip$between_2_to_4[i]!=0){
            Hip$within.2[i] <- 0
        }
    }
    
    for (i in 1:4){
        Hip[,i+2] <- as.numeric(Hip[,i+2])
    }
    
    Hip <- Hip %>% mutate(within_2.rate = within_2/total) %>%
        select(institution, within_2.rate) %>%
        t() %>% as.data.frame(stringsAsFactors = FALSE)
    
    names(Hip) <- Hip[1,]
    
    Hip <- Hip[-1,]
    
    Hip <- data.frame(lapply(Hip[1,], FUN = function(x) as.numeric(x)))
    
    
    for (i in 1:length(Hip.frame)){
        if (names(Hip.frame)[i] %in% names(Hip)){
            Hip.frame[i] <- Hip[names(Hip.frame)[i]]
        }
    }
    
    # Replace NA with N.A. for use in Excel
    
    # Hip.prod <- lapply(Hip.frame, function(x){ifelse(is.na(x), "N.A.", x)})
    # Hip.prod <- data.frame(Hip.prod)
    
    # Return production dataframe
    
    Hip.prod <- Hip.frame
    
    Hip.prod
    
}

# kpi.5  --------------------------------------------------------

kpi.5.1.html <- function(Mmm){
    
    # Announce fx started
    
    message("[kpi.5.1] Function started ", Sys.time())
    
    kpi_source_helper(Mmm)
    
    path <- grep("(.*kpi.8 .*html)", source.kpi$filepath, value = TRUE)
    
    DS_SDS.frame <- empty.frame
    DS_SDS.frame[1,] <- NA
    
    DS_SDS <- read_Chtml(path, 1)
    
    DS_SDS <- DS_SDS[,1:6]
    
    names(DS_SDS) <- c("cluster", "institution", "specialty", "total", "DS", "SDS")
    
    DS_SDS$institution <- gsub("^Subtotal.*", "Subtotal", DS_SDS$institution)
    DS_SDS$cluster <- gsub("^Grand.*", "HA", DS_SDS$cluster)
    DS_SDS$institution <- gsub("^Grand.*", "HA", DS_SDS$institution)
    DS_SDS$specialty <- gsub("^Subtotal.*", NA, DS_SDS$specialty)
    
    DS_SDS <- DS_SDS %>% filter(cluster == "KW" | cluster == "HA" | institution == "Subtotal")
    
    DS_SDS$institution <- gsub(".*total.*", NA, DS_SDS$institution)
    
    DS_SDS$cluster <- gsub(" $", "", DS_SDS$cluster)
    DS_SDS$cluster <- sapply(DS_SDS$cluster, FUN = function(x){ifelse(x=="HA", x, paste(x,"C", sep = ""))})
    
    
    for (i in 1:nrow(DS_SDS)){
        if (is.na(DS_SDS$institution[i])){
            DS_SDS$institution[i] <- DS_SDS$cluster[i]
        }
        if (is.na(DS_SDS$DS[i]) & DS_SDS$total[i]!=0){
            DS_SDS$DS[i] <- 0
        }
        if (is.na(DS_SDS$SDS[i]) & DS_SDS$total[i]!=0){
            DS_SDS$SDS[i] <- 0
        }
    }
    
    for (i in 1:3){
        DS_SDS[,i+3] <- as.numeric(DS_SDS[,i+3])
    }
    
    
    DS_SDS <- DS_SDS %>% filter(cluster=="KWC" | cluster=="HA") %>% 
        group_by(institution) %>% summarise(DS = sum(DS), SDS = sum(SDS), total = sum(total))
    
    DS_SDS <- DS_SDS %>% mutate(DS_SDS = DS + SDS) %>%
        mutate(DS_SDS.rate = DS_SDS/total) %>%
        select(institution, DS_SDS.rate) %>%
        t() %>% as.data.frame(stringsAsFactors = FALSE)
    
    names(DS_SDS) <- DS_SDS[1,]
    
    DS_SDS <- DS_SDS[-1,]
    
    DS_SDS <- data.frame(lapply(DS_SDS[1,], FUN = function(x) as.numeric(x)))
    
    
    for (i in 1:length(DS_SDS.frame)){
        if (names(DS_SDS.frame)[i] %in% names(DS_SDS)){
            DS_SDS.frame[i] <- DS_SDS[names(DS_SDS.frame)[i]]
        }
    }
    
    # Replace NA with N.A. for use in Excel
    
    # DS_SDS.prod <- lapply(DS_SDS.frame, function(x){ifelse(is.na(x), "N.A.", x)})
    # DS_SDS.prod <- data.frame(DS_SDS.prod)
    
    # Return production dataframe
    
    DS_SDS.prod <- DS_SDS.frame
    
    DS_SDS.prod
}

# kpi.8 ----------------------------------------

# kpi.8 Rate of Day Surgery plus Same Day Surgery for different specialty -----------

kpi.8.html <- function(Mmm, spec = "Overall", inst = "KWC"){
    
    # Announce fx started
    
    message("[kpi.8] Function started ", Sys.time())
    
    # Check input
    
    specialty.list <- c("Overall",
                        "ENT",
                        "O&G",
                        "OPH",
                        "ORT",
                        "SUR")
    
    if(!(spec %in% specialty.list)){
        print("specialty not allowed")
        return(0)
    }
    
    kpi_source_helper(Mmm)
    
    path <- grep("(.*kpi.8 .*html)", source.kpi$filepath, value = TRUE)
    
    DS_SDS.frame <- empty.frame
    DS_SDS.frame[1,] <- NA
    
    DS_SDS <- read_Chtml(path, 1)
    
    DS_SDS <- DS_SDS[,1:6]
    
    names(DS_SDS) <- c("cluster", "institution", "specialty", "total", "DS", "SDS")
    
    DS_SDS$institution <- gsub("^Subtotal.*", "Subtotal", DS_SDS$institution)
    DS_SDS$cluster <- gsub("^Grand.*", "HA", DS_SDS$cluster)
    DS_SDS$institution <- gsub("^Grand.*", "HA", DS_SDS$institution)
    DS_SDS$specialty <- gsub("^Subtotal.*", NA, DS_SDS$specialty)
    DS_SDS$specialty <- gsub("^Grand.*", NA, DS_SDS$specialty)
    
    if(spec == "Overall"){
        DS_SDS <- DS_SDS %>% filter(cluster == "KW" | cluster == "HA" | institution == "Subtotal")
        
        DS_SDS$institution <- gsub(".*total.*", NA, DS_SDS$institution)
        
        DS_SDS$cluster <- gsub(" $", "", DS_SDS$cluster)
        DS_SDS$cluster <- sapply(DS_SDS$cluster, FUN = function(x){ifelse(x=="HA", x, paste(x,"C", sep = ""))})
        
        
        for (i in 1:nrow(DS_SDS)){
            if (is.na(DS_SDS$institution[i])){
                DS_SDS$institution[i] <- DS_SDS$cluster[i]
            }
            if (is.na(DS_SDS$DS[i]) & DS_SDS$total[i]!=0){
                DS_SDS$DS[i] <- 0
            }
            if (is.na(DS_SDS$SDS[i]) & DS_SDS$total[i]!=0){
                DS_SDS$SDS[i] <- 0
            }
        }
        
        for (i in 1:3){
            for (j in 1:nrow(DS_SDS)){
                if(is.na(DS_SDS[j,i+3])){
                    DS_SDS[j,i+3] <- "0"
                }
            }
            DS_SDS[,i+3] <- as.numeric(DS_SDS[,i+3])
        }
        
        DS_SDS <- DS_SDS %>% #filter(cluster=="KWC" | cluster=="HA") %>% 
            group_by(institution) %>% summarise(SDS = sum(SDS), DS = sum(DS), total = sum(total)) %>%
            t() %>% as.data.frame(stringsAsFactors = FALSE)
        
        names(DS_SDS) <- DS_SDS[1,]
        
        DS_SDS <- DS_SDS[-1,]
        
        for (i in 1:length(DS_SDS)){
            DS_SDS[,i] <- sapply(DS_SDS[,i], FUN = function(x) as.numeric(x))
        }
        
        if(inst == "KWC"){
            DS_SDS.prod <- DS_SDS %>% select(CMC, KWH, NLTH, OLM, PMH, YCH, KWC, HA)
        }else if(inst == "HA"){
            DS_SDS.prod <- DS_SDS %>% select(HKEC, HKWC, KCC, KEC, NTEC, NTWC, KWC, HA)
        }
        
    }else{
        
        DS_SDS <- DS_SDS %>% filter(specialty==spec)
        DS_SDS$cluster <- sapply(DS_SDS$cluster, FUN = function(x){ifelse(x=="HA", x, paste(x,"C", sep = ""))})
        
        for (i in 1:3){
            for (j in 1:nrow(DS_SDS)){
                if(is.na(DS_SDS[j,i+3])){
                    DS_SDS[j,i+3] <- "0"
                }
            }
            DS_SDS[,i+3] <- as.numeric(DS_SDS[,i+3])
        }
        
        # Cluster based
        
        DS_SDS.HA <- DS_SDS %>% group_by(cluster) %>% arrange(cluster) %>% summarise(SDS = sum(SDS), DS = sum(DS), total = sum(total)) %>%
            t() %>% as.data.frame(stringsAsFactors = FALSE)
        
        names(DS_SDS.HA) <- DS_SDS.HA[1,]
        
        DS_SDS.HA <- DS_SDS.HA[-1,]
        
        for (k in 1:ncol(DS_SDS.HA)){
            DS_SDS.HA[,k] <- as.numeric(DS_SDS.HA[,k])
        }
        
        DS_SDS.HA <- DS_SDS.HA %>% mutate(KWC_ = KWC) %>% select(-KWC) %>% mutate(HA = rowSums(.))
        
        names(DS_SDS.HA)[which(names(DS_SDS.HA)=="KWC_")] <- "KWC"
        
        # Cluster based
        
        DS_SDS.KWC <- DS_SDS %>% filter(cluster=="KWC") %>% arrange(institution) %>% select(institution, SDS, DS, total) %>%
            t() %>% as.data.frame(stringsAsFactors = FALSE)
        
        names(DS_SDS.KWC) <- DS_SDS.KWC[1,]
        
        DS_SDS.KWC <- DS_SDS.KWC[-1,]
        
        for (k in 1:ncol(DS_SDS.KWC)){
            DS_SDS.KWC[,k] <- as.numeric(DS_SDS.KWC[,k])
        }
        
        h <- ncol(DS_SDS.HA)
        
        DS_SDS.KWC <- cbind(DS_SDS.KWC, DS_SDS.HA[,c(h-1,h)])
        
        names(DS_SDS.KWC)[which(names(DS_SDS.KWC)=="KWC_")] <- "KWC"
        
        if(inst=="KWC"){
            DS_SDS.prod <- DS_SDS.KWC
        }else if(inst=="HA"){
            DS_SDS.prod <- DS_SDS.HA
        }
    }
    
    # Return production dataframe
    
    DS_SDS.prod
}

