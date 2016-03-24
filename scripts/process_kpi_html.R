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


kpi.3.3.html <- function(Mmm, trend = FALSE, row){
    
    # Announce fx started
    
    message("[kpi.3.3] Function started ", Sys.time())
    
    require("tidyr")
    
    kpi_source_helper(Mmm)
    
    path <- grep("(.*kpi.3.3 .*html)", source.kpi$filepath, value = TRUE)
    
    MRSA.frame <- empty.frame
    
    lines <- readLines(path)
    MRSA <- suppressWarnings(htmltab(doc = lines, which = 3, header = 1))

    MRSA <- MRSA %>% select(Period, Hospital, contains("acute"))
    
    names(MRSA) <- gsub("(^No. of )", "", names(MRSA))
    names(MRSA) <- gsub("(^.*/.*$)", "per_1000_PD", names(MRSA))
    names(MRSA) <- gsub("( )", "_", names(MRSA))
    names(MRSA) <- gsub("(_Bacteremia)", "", names(MRSA))
    
    
    for (i in 1:3){
        MRSA[,i+2] <- as.numeric(MRSA[,i+2])
    }
    
    MRSA <- MRSA %>% mutate(per_1000_PD = MRSA_in_Acute_Beds/Acute_Patient_Days*1000)
    
    MRSA.HA <- MRSA %>% group_by(Period) %>%
        summarise(Hospital = "HA", MRSA_in_Acute_Beds = sum(MRSA_in_Acute_Beds), Acute_Patient_Days = sum(Acute_Patient_Days)) %>%
        mutate(per_1000_PD = MRSA_in_Acute_Beds/Acute_Patient_Days*1000)
    
    MRSA <- MRSA %>% filter(Hospital %in% c(names(MRSA.frame), "NLT")) %>% arrange(Period, Hospital)
    
    MRSA.KWC <- MRSA %>% group_by(Period) %>%
        summarise(Hospital = "KWC", MRSA_in_Acute_Beds = sum(MRSA_in_Acute_Beds), Acute_Patient_Days = sum(Acute_Patient_Days)) %>%
        mutate(per_1000_PD = MRSA_in_Acute_Beds/Acute_Patient_Days*1000)
    
    MRSA <- bind_rows(MRSA, MRSA.KWC, MRSA.HA)
    
    if(trend==TRUE){
        
        MRSA.req <- MRSA %>% select(-MRSA_in_Acute_Beds, -Acute_Patient_Days) %>% spread(key = Period, value = per_1000_PD)
        
        MRSA.KWC.req <- MRSA.req %>% filter(Hospital=="KWC")
        MRSA.HA.req <- MRSA.req %>% filter(Hospital=="HA")
        MRSA.req <- MRSA.req %>% filter(Hospital!="HA" & Hospital!="KWC") #%>%
        #filter(Hospital!="KCH" & Hospital!="WTS")
        
        
        MRSA.req <- bind_rows(MRSA.req, MRSA.KWC.req, MRSA.HA.req)
        
        rownames(MRSA.req) <- MRSA.req$Hospital
        
        MRSA.req <- MRSA.req %>% select(-Hospital)
        
        # Replace NA with ""
        
        MRSA.req[is.na(MRSA.req)] <- ""
        
        MRSA.prod <- MRSA.req
        
    }else{
        
        MRSA.req <- MRSA %>% ungroup() %>% group_by(Hospital) %>%
            summarise(MRSA_in_Acute_Beds = sum(MRSA_in_Acute_Beds), Acute_Patient_Days = sum(Acute_Patient_Days)) %>%
            mutate(per_1000_PD = MRSA_in_Acute_Beds/Acute_Patient_Days*1000) %>%
            select(-MRSA_in_Acute_Beds, -Acute_Patient_Days) %>%
            t() %>% as.data.frame(stringsAsFactors=FALSE)
        
        names(MRSA.req) <- MRSA.req[1,]
        MRSA.req <- MRSA.req[-1,]
        
        # Fill empty frame
        
        names(MRSA.req) <- gsub("NLT", "NLTH", names(MRSA.req))
        
        MRSA.frame[1,] <- NA
        
        for (i in 1:length(MRSA.frame)){
            if (names(MRSA.frame)[i] %in% names(MRSA.req)){
                MRSA.frame[i] <- MRSA.req[names(MRSA.frame)[i]]
            }
        }
        
        # Replace NA with N.A. for Excel use
        
        MRSA.prod <- data.frame(lapply(MRSA.frame, FUN = function(x){ifelse(is.na(x), "N.A.", x)}), stringsAsFactors = FALSE)
        
        for (i in 1:ncol(MRSA.prod)){
            if ("N.A." %in% MRSA.prod[,i]){
            } else {
                MRSA.prod[,i] <- as.numeric(MRSA.prod[,i])
            }
        }
        
    }
    
    # Return production ready dataframe
    
    MRSA.prod
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

# tre.3 ---------------------------------------

tre.3.html <- function(Mmm, MED, write_db = FALSE, backup = FALSE){
    
    # Announce fx started
    
    message("[tre.3] Function started ", Sys.time())
    
    require(reshape2)
    
    kpi_source_helper(Mmm)
    
    if(MED==FALSE){
        path.Atn <- grep("(.*tre.3.1.1 .*html)", source.kpi$filepath, value = TRUE)
        path.Adm <- grep("(.*tre.3.1.2 .*html)", source.kpi$filepath, value = TRUE)
        
        ATN_T <- read_Chtml(path.Atn, 1:2)
        ADM_T <- read_Chtml(path.Adm, 1:2)
        
    }else if(MED==TRUE){
        path.Atn <- grep("(.*tre.3.2.1 .*html)", source.kpi$filepath, value = TRUE)
        path.Adm <- grep("(.*tre.3.2.2 .*html)", source.kpi$filepath, value = TRUE)
        
        ATN_T <- read_Chtml(path.Atn, 1:2)
        ADM_T <- read_Chtml(path.Adm, 1:2)
    }
    
    # Name columns
    
    names(ATN_T)[1] <- "Age"
    names(ADM_T)[1] <- "Age"
    
    names(ATN_T)[2] <- "Sex"
    names(ADM_T)[2] <- "Sex"
    
    names(ATN_T)[3] <- "Amb"
    names(ADM_T)[3] <- "Amb"
    
    names(ATN_T)[4] <- "Tri"
    names(ADM_T)[4] <- "Tri"
    
    names(ATN_T)[ncol(ATN_T)] <- "HA"
    names(ADM_T)[ncol(ADM_T)] <- "HA"
    
    names(ATN_T) <- gsub(" >> No. of A&E 1st Attendances", "", names(ATN_T))
    names(ADM_T) <- gsub(" >> No. of A&E 1st Attendances", "", names(ADM_T))
    
    # Filter unused rows

    ATN_T <- ATN_T %>% filter(Sex!="Grand Total :")
    ADM_T <- ADM_T %>% filter(Sex!="Grand Total :")
    
    # Use numbers to code groups
    
    ## Age
    if(MED==FALSE){
        ATN_T$Age <- gsub("^0 .*", 1, ATN_T$Age)
        ATN_T$Age <- gsub("^5 .*", 2, ATN_T$Age)
        ATN_T$Age <- gsub("^15 .*", 3, ATN_T$Age)
        ATN_T$Age <- gsub("^35 .*", 4, ATN_T$Age)
        ATN_T$Age <- gsub("^55 .*", 5, ATN_T$Age)
        ATN_T$Age <- gsub("^65 .*", 6, ATN_T$Age)
        ATN_T$Age <- gsub("^75 .*", 7, ATN_T$Age)
        ATN_T$Age <- gsub("^85.*", 8, ATN_T$Age)
        
        ADM_T$Age <- gsub("^0 .*", 1, ADM_T$Age)
        ADM_T$Age <- gsub("^5 .*", 2, ADM_T$Age)
        ADM_T$Age <- gsub("^15 .*", 3, ADM_T$Age)
        ADM_T$Age <- gsub("^35 .*", 4, ADM_T$Age)
        ADM_T$Age <- gsub("^55 .*", 5, ADM_T$Age)
        ADM_T$Age <- gsub("^65 .*", 6, ADM_T$Age)
        ADM_T$Age <- gsub("^75 .*", 7, ADM_T$Age)
        ADM_T$Age <- gsub("^85.*", 8, ADM_T$Age)
    }else if(MED==TRUE){
        ATN_T$Age <- gsub("^18 .*", 1, ATN_T$Age)
        ATN_T$Age <- gsub("^35 .*", 2, ATN_T$Age)
        ATN_T$Age <- gsub("^55 .*", 3, ATN_T$Age)
        ATN_T$Age <- gsub("^65 .*", 4, ATN_T$Age)
        ATN_T$Age <- gsub("^75 .*", 5, ATN_T$Age)
        ATN_T$Age <- gsub("^85.*", 6, ATN_T$Age)
        
        ADM_T$Age <- gsub("^18 .*", 1, ADM_T$Age)
        ADM_T$Age <- gsub("^35 .*", 2, ADM_T$Age)
        ADM_T$Age <- gsub("^55 .*", 3, ADM_T$Age)
        ADM_T$Age <- gsub("^65 .*", 4, ADM_T$Age)
        ADM_T$Age <- gsub("^75 .*", 5, ADM_T$Age)
        ADM_T$Age <- gsub("^85.*", 6, ADM_T$Age)
    }
    
    
    ## Sex
    ATN_T$Sex <- gsub("F", 1, ATN_T$Sex)
    ATN_T$Sex <- gsub("M", 2, ATN_T$Sex)
    
    ADM_T$Sex <- gsub("F", 1, ADM_T$Sex)
    ADM_T$Sex <- gsub("M", 2, ADM_T$Sex)
    
    ## Amblatory Case
    ATN_T$Amb <- gsub("N", 1, ATN_T$Amb)
    ATN_T$Amb <- gsub("Y", 2, ATN_T$Amb)
    
    ADM_T$Amb <- gsub("N", 1, ADM_T$Amb)
    ADM_T$Amb <- gsub("Y", 2, ADM_T$Amb)
    
    ## Triage
    ATN_T$Tri <- gsub(" .*$", "", ATN_T$Tri)
    ADM_T$Tri <- gsub(" .*$", "", ADM_T$Tri)
    
    
    
    # Select relevant columns
    
    ATN_T <- ATN_T %>% select(Age, Sex, Amb, Tri, CMC, KWH, NLTH, PMH, YCH, HA)
    
    ADM_T <- ADM_T %>% select(Age, Sex, Amb, Tri, CMC, KWH, NLTH, PMH, YCH, HA)
    
    # Convert to numeric format
    
    for (i in 1:ncol(ATN_T)){ # As No of 1st attenance is alwasy > than No of Admissions)
        ATN_T[,i] <- sapply(ATN_T[,i], as.numeric)
        ADM_T[,i] <- sapply(ADM_T[,i], as.numeric)
    }
    
    
    # Calculate cluster sum
    
    ## Calculate Attendance
    ATN_T$KWC <- NA
    
    for (i in 1:nrow(ATN_T)){
        ATN_T$KWC[i] <- sum(ATN_T[i, 5:(ncol(ATN_T)-2)], na.rm = TRUE)
    }
    ## Calculate Admission
    ADM_T$KWC <- NA
    
    for (i in 1:nrow(ADM_T)){
        ADM_T$KWC[i] <- sum(ADM_T[i, 5:(ncol(ADM_T)-2)], na.rm = TRUE)
    }
    
    
    # Check missing rows
    
    ## Check Attendance
    if(MED==FALSE){
        x <- 8
    }else if(MED==TRUE){
        x <- 6
    }
    
    for (i in 1:x){ # 8/6 age groups
        for (j in 1:2){ # 2 sex groups
            for (k in 1:2){ # 2 ambulatory care groups
                for (t in 1:5){ # 5 triage groups
                    
                    count <- ATN_T %>% filter(Age==i, Sex==j, Amb==k, Tri==t)
                    
                    if (nrow(count) != 1){
                        temp_df <- data.frame(i, j, k, t, NA, NA, NA, NA, NA, NA, NA)
                        names(temp_df) <- names(ATN_T)
                        ATN_T <- bind_rows(ATN_T, temp_df)
                    }
                    
                }
            }
        }
    }
    ## Check Admission
    for (i in 1:x){ # 8/6 age groups
        for (j in 1:2){ # 2 sex groups
            for (k in 1:2){ # 2 ambulatory care groups
                for (t in 1:5){ # 5 triage groups
                    
                    count <- ADM_T %>% filter(Age==i, Sex==j, Amb==k, Tri==t)
                    
                    if (nrow(count) != 1){
                        temp_df <- data.frame(i, j, k, t, NA, NA, NA, NA, NA, NA, NA)
                        names(temp_df) <- names(ADM_T)
                        ADM_T <- bind_rows(ADM_T, temp_df)
                    }
                    
                }
            }
        }
    }
    
    # Melt into narrow format
    
    ATN_Tr <- melt(ATN_T, id=c("Age", "Sex", "Amb", "Tri")) %>% arrange(Age, Sex, Amb, Tri, variable)
    ADM_Tr <- melt(ADM_T, id=c("Age", "Sex", "Amb", "Tri")) %>% arrange(Age, Sex, Amb, Tri, variable)
    
    # Name columns
    
    names(ATN_Tr)[5] <- "Hosp"
    names(ATN_Tr)[6] <- "Attndance"
    names(ADM_Tr)[6] <- "Admission"
    
    # Merge into single dataframe
    
    RATE <- bind_cols(ATN_Tr, ADM_Tr[6])
    
    RATE$Attndance[is.na(RATE$Attndance)] <- 0
    RATE$Admission[is.na(RATE$Admission)] <- 0
    
    # Order Hosp with code
    
    RATE$Hosp <- as.character(RATE$Hosp)
    
    RATE$Hosp[which(RATE$Hosp=="CMC")] <- 1
    RATE$Hosp[which(RATE$Hosp=="KWH")] <- 2
    RATE$Hosp[which(RATE$Hosp=="NLTH")] <- 3
    RATE$Hosp[which(RATE$Hosp=="PMH")] <- 4
    RATE$Hosp[which(RATE$Hosp=="YCH")] <- 5
    RATE$Hosp[which(RATE$Hosp=="KWC")] <- 6
    RATE$Hosp[which(RATE$Hosp=="HA")] <- 7
    
    RATE$Hosp <- as.numeric(RATE$Hosp)
    
    RATE <- RATE %>% arrange(Age, Sex, Amb, Tri, Hosp)
    
    RATE$CrudeRate <- NA
    
    for (i in 1:nrow(RATE)){
        if (RATE$Hosp[i]==7){
            RATE$CrudeRate[i] <- RATE$Admission[i]/RATE$Attndance[i]
        }
    }
    
    RATE$CrudeRate[is.na(RATE$CrudeRate)] <- 0
    
    # Assign crude rate to all rows
    
    for (i in 0:(x*20-1)){
        
        si <- 1+7*i
        ei <- 6+7*i
        ri <- 7+7*i
        
        RATE$CrudeRate[si:ei] <- RATE$CrudeRate[ri]
    }
    
    # Calculate expected admissions
    
    RATE <- RATE %>% mutate(ExpectedAdm = Attndance * CrudeRate)
    
    # Subset for each hosp
    
    R_HA <- RATE %>% filter(Hosp==7) %>% select(Attndance, Admission) %>%
        summarise(Atn=sum(Attndance), Adm=sum(Admission)) %>% mutate(AdmRate=Adm/Atn)
    adm_rate <- as.numeric(R_HA$AdmRate) # Overall admission rate
    
    R_H1 <- RATE %>% filter(Hosp==1) %>% select(Attndance, Admission, ExpectedAdm) %>%
        summarise(Atn=sum(Attndance), Adm=sum(Admission), Exp=sum(ExpectedAdm)) %>%
        mutate(Rate=Adm/Exp)
    H1_rate <- as.numeric(R_H1$Rate*adm_rate) # Standardised admission rate
    
    R_H2 <- RATE %>% filter(Hosp==2) %>% select(Attndance, Admission, ExpectedAdm) %>%
        summarise(Atn=sum(Attndance), Adm=sum(Admission), Exp=sum(ExpectedAdm)) %>%
        mutate(Rate=Adm/Exp)
    H2_rate <- as.numeric(R_H2$Rate*adm_rate) # Standardised admission rate 
    
    R_H3 <- RATE %>% filter(Hosp==3) %>% select(Attndance, Admission, ExpectedAdm) %>%
        summarise(Atn=sum(Attndance), Adm=sum(Admission), Exp=sum(ExpectedAdm)) %>%
        mutate(Rate=Adm/Exp)
    H3_rate <- as.numeric(R_H3$Rate*adm_rate) # Standardised admission rate
    
    R_H4 <- RATE %>% filter(Hosp==4) %>% select(Attndance, Admission, ExpectedAdm) %>%
        summarise(Atn=sum(Attndance), Adm=sum(Admission), Exp=sum(ExpectedAdm)) %>%
        mutate(Rate=Adm/Exp)
    H4_rate <- as.numeric(R_H4$Rate*adm_rate) # Standardised admission rate
    
    R_H5 <- RATE %>% filter(Hosp==5) %>% select(Attndance, Admission, ExpectedAdm) %>%
        summarise(Atn=sum(Attndance), Adm=sum(Admission), Exp=sum(ExpectedAdm)) %>%
        mutate(Rate=Adm/Exp)
    H5_rate <- as.numeric(R_H5$Rate*adm_rate) # Standardised admission rate
    
    R_C <- RATE %>% filter(Hosp==6) %>% select(Attndance, Admission, ExpectedAdm) %>%
        summarise(Atn=sum(Attndance), Adm=sum(Admission), Exp=sum(ExpectedAdm)) %>%
        mutate(Rate=Adm/Exp)
    HC_rate <- as.numeric(R_C$Rate*adm_rate) # Standardised admission rate
    
    RATE_Tr <- data.frame(
        c(H1_rate, NA, H2_rate, H3_rate, NA, H4_rate, NA, H5_rate, HC_rate, adm_rate)
    )
    
    names(RATE_Tr) <- Series[1,12]
    
    # Move Inst to row names
    
    row.names(RATE_Tr) <- c("CMC", "KCH", "KWH", "NLTH", "OLMH", "PMH", "WTSH", "YCH", "KWC", "HA")
    
    # Store result in database
    
    if(MED==FALSE){
        
        db <- read.csv(file.path("source", "db", "tre.3.1.db.csv"), stringsAsFactors = FALSE)
        tmp_series <- gsub(" ", ".", Series[1,])
        
        tmp_df <- db[names(db) %in% tmp_series]
        
        row.names(tmp_df) <- row.names(RATE_Tr)
        
        if(!(tmp_series[12] %in% names(db))){
            tmp_df <- bind_cols(tmp_df, RATE_Tr)
            row.names(tmp_df) <- row.names(RATE_Tr)
            
            if(write_db==TRUE){
                # db backup
                if(backup==TRUE){
                    dir.create(file.path("source", "db", "backup", Dates[1,]), showWarnings = FALSE)
                    file.copy(from = file.path("source", "db", "tre.3.1.db.csv"),
                              to = file.path("source", "db", "backup", Dates[1,], "tre.3.1.db.csv"),
                              overwrite = TRUE)    
                }
                
                db <- bind_cols(db, RATE_Tr)
                write.csv(db, file.path("source", "db", "tre.3.1.db.csv"), row.names = FALSE)
            }
            
        }else{
            tmp_df[,which(names(tmp_df)==tmp_series[12])] <- RATE_Tr[,1]
            
            if(write_db==TRUE){
                # db backup
                if(backup==TRUE){
                    dir.create(file.path("source", "db", "backup", Dates[1,]), showWarnings = FALSE)
                    file.copy(from = file.path("source", "db", "tre.3.1.db.csv"),
                              to = file.path("source", "db", "backup", Dates[1,], "tre.3.1.db.csv"),
                              overwrite = TRUE)    
                }
                
                db[,which(names(db)==tmp_series[12])] <- RATE_Tr[,1]
                write.csv(db, file.path("source", "db", "tre.3.1.db.csv"), row.names = FALSE)
            }
            
        }
        
        
    }else if(MED==TRUE){
        
        db <- read.csv(file.path("source", "db", "tre.3.2.db.csv"), stringsAsFactors = FALSE)
        tmp_series <- gsub(" ", ".", Series[1,])
        
        tmp_df <- db[names(db) %in% tmp_series]
        
        row.names(tmp_df) <- row.names(RATE_Tr)
        
        if(!(tmp_series[12] %in% names(db))){
            tmp_df <- bind_cols(tmp_df, RATE_Tr)
            row.names(tmp_df) <- row.names(RATE_Tr)
            
            if(write_db==TRUE){
                # db backup
                if(backup==TRUE){
                    dir.create(file.path("source", "db", "backup", Dates[1,]), showWarnings = FALSE)
                    file.copy(from = file.path("source", "db", "tre.3.2.db.csv"),
                              to = file.path("source", "db", "backup", Dates[1,], "tre.3.2.db.csv"),
                              overwrite = TRUE)    
                }
                
                db <- bind_cols(db, RATE_Tr)
                write.csv(db, file.path("source", "db", "tre.3.2.db.csv"), row.names = FALSE)
            }
            
        }else{
            tmp_df[,which(names(tmp_df)==tmp_series[12])] <- RATE_Tr[,1]
            
            if(write_db==TRUE){
                # db backup
                if(backup==TRUE){
                    dir.create(file.path("source", "db", "backup", Dates[1,]), showWarnings = FALSE)
                    file.copy(from = file.path("source", "db", "tre.3.2.db.csv"),
                              to = file.path("source", "db", "backup", Dates[1,], "tre.3.2.db.csv"),
                              overwrite = TRUE)    
                }
                
                db[,which(names(db)==tmp_series[12])] <- RATE_Tr[,1]
                write.csv(db, file.path("source", "db", "tre.3.2.db.csv"), row.names = FALSE)
            }
            
        }
        
    }
    
    # Remove NA rows
    
    tmp_df <- tmp_df %>% mutate(Inst=row.names(tmp_df)) %>% filter(!(row.names(tmp_df) %in% c("KCH", "OLMH", "WTSH")))
    
    row.names(tmp_df) <- tmp_df$Inst
    
    tmp_df <- tmp_df %>% select(-Inst)
    
    # Return dataframe
    
    tmp_df
    
}

# tre.6 Stroke --------------------------------------------------------------

tre.6.html <- function(Mmm){
    
    # Announce fx started
    
    message("[tre.6] Function started ", Sys.time())
    
    kpi_source_helper(Mmm)
    
    path <- grep("(.*tre.6 .*html)", source.kpi$filepath, value = TRUE)
    
    
    ASU_T <- read_Chtml(path, 1:3)
    
    index <- c(1, 2, 3, 3+11, 3+12, ncol(ASU_T)-1)
    
    
    # Name columns
    
    names(ASU_T) <- gsub(" >> No. of Episodes", "", names(ASU_T))
    
    names(ASU_T)[1] <- "Cluster"
    names(ASU_T)[2] <- "Inst"
    names(ASU_T)[ncol(ASU_T)] <- "RowSum"
    
    # Tidy names
    
    for(i in 1:nrow(ASU_T)){
        if(grepl("Subtotal", ASU_T[i,2])){
            ASU_T[i,2] <- ASU_T[i,1]
        }
    }
    
    ASU_T[nrow(ASU_T),1] <- "HA"
    ASU_T[nrow(ASU_T),2] <- "HA"
    
    ASU_T$Cluster <- gsub(" ", "", ASU_T$Cluster)
    
    # Filter unused rows
    
    ASU_T <- ASU_T %>% filter(Cluster=="KW"|Cluster=="HA")
    ASU_T$Cluster <- 1:nrow(ASU_T)
    names(ASU_T)[1] <- "Order"
    
    # Melt dataframe
    
    ASU_Tr <- ASU_T %>% select(index[1], index[2], index[3]:index[6]) %>% melt(id=c("Order", "Inst"))
    
    
    
    # Add Month variable
    
    ASU_Tr$Month <- gsub(".*([0-9]{4}-[0-9]{2})", "\\1", ASU_Tr$variable)
    ASU_Tr$value <-sapply(ASU_Tr$value, function(x)as.numeric(x))
    ASU_Tr$ASU <- sapply(ASU_Tr$variable, function(x)if(grepl("ASU", x)){"Treated"}else{"NotTreated"})
    
    
    ASU_Tr <- ASU_Tr %>% select(Order, Inst, Month, ASU, value) %>% dcast(Order + Inst + Month ~ ASU) %>% mutate(Total=Treated + NotTreated)
    
    ## Logic to handle NA values
    
    for (i in 1:nrow(ASU_Tr)){
        
        if(is.na(ASU_Tr$Treated[i])){
            if(!is.na(ASU_Tr$NotTreated[i])){
                ASU_Tr$Treated[i] <- 0
            }
        }
        
        
        if(is.na(ASU_Tr$Total[i])){
            if(!is.na(ASU_Tr$NotTreated[i])|!is.na(ASU_Tr$Treated[i])){
                ASU_Tr$Total[i] <- sum(ASU_Tr$NotTreated[i], ASU_Tr$Treated[i], na.rm = TRUE)
            }
        }
    }
    
    # Calculate % treated in ASU
    
    ASU_Tr <- ASU_Tr %>% mutate(PercentTreated = Treated/Total) %>% select(Order, Inst, Month, PercentTreated)
    
    # Filter NON-KWC inst
    
    ASU_Tr <- dcast(ASU_Tr, Order + Inst ~ Month, value.var = "PercentTreated")
    
    # Move Inst to row names
    
    row.names(ASU_Tr) <- ASU_Tr$Inst
    
    # Remove excess columns
    
    ASU_Tr <- ASU_Tr %>% select(-Order, -Inst)
    
    
    # Return dataframe
    
    ASU_Tr
    
}

# tre.7 Fracture Hip -------------------------------------------------------

tre.7.html <- function(Mmm){
    
    # Announce fx started
    
    message("[tre.7] Function started ", Sys.time())
    
    kpi_source_helper(Mmm)
    
    path <- grep("(.*tre.7 .*html)", source.kpi$filepath, value = TRUE)
    
    
    HIP_T <- read_Chtml(path, 1:3)
    
    index <- c(1, 2, 3, 3+11, 3+12, 3+23, 3+24, ncol(HIP_T)-1)
    
    
    # Name columns
    
    names(HIP_T) <- gsub(" >> No. of Episodes", "", names(HIP_T))
    
    names(HIP_T)[1] <- "Cluster"
    names(HIP_T)[2] <- "Inst"
    names(HIP_T)[ncol(HIP_T)] <- "RowSum"
    
    # Tidy names
    
    for(i in 1:nrow(HIP_T)){
        if(grepl("Subtotal", HIP_T[i,2])){
            HIP_T[i,2] <- HIP_T[i,1]
        }
    }
    
    HIP_T[nrow(HIP_T),1] <- "HA"
    HIP_T[nrow(HIP_T),2] <- "HA"
    
    HIP_T$Cluster <- gsub(" ", "", HIP_T$Cluster)
    
    # Filter unused rows
    
    HIP_T <- HIP_T %>% filter(Cluster=="KW"|Cluster=="HA")
    HIP_T$Cluster <- 1:nrow(HIP_T)
    names(HIP_T)[1] <- "Order"
    
    # Melt dataframe
    
    HIP_Tr <- HIP_T %>% select(index[1], index[2], index[3]:index[4], index[5]:index[6], index[7]:index[8]) %>% melt(id=c("Order", "Inst"))
    
    
    # Add Month variable
    
    HIP_Tr$Month <- gsub(".* >> ([0-9]{4}-[0-9]{2})", "\\1", HIP_Tr$variable)

    HIP_Tr$value <-sapply(HIP_Tr$value, function(x)as.numeric(x))
    HIP_Tr$Time <- sapply(HIP_Tr$variable, function(x)if(grepl("within", x)){"Within"}else{"Outside"})
    
    
    HIP_Tr <- HIP_Tr %>% select(Order, Inst, Month, Time, value) %>% dcast(Order + Inst + Month ~ Time, value.var = "value", fun.aggregate = sum, na.rm = TRUE) %>% mutate(Total=Within + Outside)
    
    ## Logic to handle NA values
    
    for (i in 1:nrow(HIP_Tr)){
        
        if(is.na(HIP_Tr$Within[i])){
            if(!is.na(HIP_Tr$Outside[i])){
                HIP_Tr$Within[i] <- 0
            }
        }
        
        if(is.na(HIP_Tr$Outside[i])){
            if(!is.na(HIP_Tr$Within[i])){
                HIP_Tr$Outside[i] <- 0
            }
        }
        
        if(is.na(HIP_Tr$Total[i])){
            if(!is.na(HIP_Tr$Outside[i])|!is.na(HIP_Tr$Within[i])){
                HIP_Tr$Total[i] <- sum(HIP_Tr$Outside[i], HIP_Tr$Within[i], na.rm = TRUE)
            }
        }
    }
    
    # Calculate % treated in ASU
    
    HIP_Tr <- HIP_Tr %>% mutate(PercentWithin = Within/Total) %>% select(Order, Inst, Month, PercentWithin)
    
    # Filter NON-KWC inst
    
    HIP_Tr <- dcast(HIP_Tr, Order + Inst ~ Month, value.var = "PercentWithin")
    
    # Move Inst to row names
    
    row.names(HIP_Tr) <- HIP_Tr$Inst
    
    # Remove excess columns
    
    HIP_Tr <- HIP_Tr %>% select(-Order, -Inst)
    
    
    # Return dataframe
    
    HIP_Tr
    
}

# tre.9 DS +SDS ------------------------------------------------------------

tre.9.html <- function(Mmm){
    
    # Announce fx started
    
    message("[tre.9] Function started ", Sys.time())
    
    kpi_source_helper(Mmm)
    
    path <- grep("(.*tre.9 .*html)", source.kpi$filepath, value = TRUE)
    
    
    SDS_T <- read_Chtml(path, 1:2)

    # Name columns

    names(SDS_T)[1] <- "Cluster"
    names(SDS_T)[2] <- "Inst"
    names(SDS_T)[ncol(SDS_T)] <- "RowSum"
    
    # Tidy names
    
    for(i in 1:nrow(SDS_T)){
        if(grepl("Subtotal", SDS_T[i,2])){
            SDS_T[i,2] <- SDS_T[i,1]
        }
    }
    
    SDS_T[nrow(SDS_T),1] <- "HA"
    SDS_T[nrow(SDS_T),2] <- "HA"
    
    SDS_T$Cluster <- gsub(" ", "", SDS_T$Cluster)
    
    # Filter unused rows
    
    SDS_T <- SDS_T %>% filter(Cluster=="KW"|Cluster=="HA")
    SDS_T$Cluster <- 1:nrow(SDS_T)
    names(SDS_T)[1] <- "Order"
    
    # Move Inst to row names
    
    row.names(SDS_T) <- SDS_T$Inst
    
    # Remove excess columns
    
    SDS_Tr <- SDS_T %>% select(-Order, -Inst, -RowSum)
    
    ## Convert to numeric datatype where applicable
    
    for (i in 1:12){
        SDS_Tr[,i] <- sapply(SDS_Tr[,i], as.numeric)
        SDS_Tr[,i] <- sapply(SDS_Tr[,i], function(x)(x/100))
    }
    
    
    # Return dataframe
    
    SDS_Tr
    
}