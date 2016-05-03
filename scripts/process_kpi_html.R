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

# kpi.4 --------------------------------------------------------

# kpi.5  --------------------------------------------------------


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

