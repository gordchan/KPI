# Process KPI
#
# To extract and reformat data ready to be filled in the excel template file.
#
# Gordon CHAN
#
# Aug 2015


# Helper function to help with reading source file ----------------------------------------

# i. Initialise empty dataframe for use by each measurements
#
# ii. Generate file table according to reporting month period specified

kpi_source_helper <- function(Mmm){
    
    require("xlsx")
    require("dplyr")
    
    # Empty dataframe
    
    empty.frame <<- data.frame(CMC = numeric(0),
                              KCH = numeric(0),
                              KWH = numeric(0),
                              NLTH = numeric(0),
                              OLMH = numeric(0),
                              PMH = numeric(0),
                              WTSH = numeric(0),
                              YCH = numeric(0),
                              KWC = numeric(0),
                              HA = numeric(0))
    
    # File table
    # Clinical KPI

    source.kpi <- read.xlsx("source/KPI items.xlsx",
                            sheetName = "KPI",
                            colIndex = 4,
                            stringsAsFactors = FALSE)
    
    source.HA.kpi <- read.xlsx("source/KPI items.xlsx",
                            sheetName = "KPI",
                            colIndex = 5,
                            stringsAsFactors = FALSE)
            names(source.HA.kpi) <- "Query.Name"
    
    source.kpi <- bind_rows(source.kpi, source.HA.kpi)

        source.kpi <<- source.kpi %>% filter(!is.na(Query.Name)) %>%
            distinct(Query.Name) %>% arrange(Query.Name) %>%
            mutate(filename = paste(Query.Name, ".xlsx", sep = "")) %>%
            mutate(filepath = file.path("source", Mmm, filename))
    
    # KPI Trend
    source.tre <- read.xlsx("source/KPI items.xlsx",
                            sheetName = "KPI",
                            colIndex = 4,
                            stringsAsFactors = FALSE)
        
    source.HA.tre <- read.xlsx("source/KPI items.xlsx",
                            sheetName = "KPI",
                            colIndex = 5,
                            stringsAsFactors = FALSE)
        names(source.HA.tre) <- "Query.Name"
        
        source.tre <- bind_rows(source.tre, source.HA.tre)
        
        source.tre <<- source.tre %>% filter(!is.na(Query.Name)) %>%
            distinct(Query.Name) %>% arrange(Query.Name) %>%
            mutate(filename = paste(Query.Name, ".xlsx", sep = "")) %>%
            mutate(filepath = file.path("source", Mmm, filename))    
    
    # SOP KPI
    source.sop <- read.xlsx("source/KPI items.xlsx",
                            sheetName = "KPI",
                            colIndex = 4,
                            stringsAsFactors = FALSE)
        
    source.HA.sop <- read.xlsx("source/KPI items.xlsx",
                            sheetName = "KPI",
                            colIndex = 5,
                            stringsAsFactors = FALSE)
    names(source.HA.sop) <- "Query.Name"
        
        source.sop <- bind_rows(source.sop, source.HA.sop)
        
        source.sop <<- source.sop %>% filter(!is.na(Query.Name)) %>%
            distinct(Query.Name) %>% arrange(Query.Name) %>%
            mutate(filename = paste(Query.Name, ".xlsx", sep = "")) %>%
            mutate(filepath = file.path("source", Mmm, filename))    
}


# KPI targets -------------------------------------------------------------

# Return stored KPI targets according to the input kpi. s/n value
#
# Accept input of kpi. s/n (i.e. "kpi.1" to "kpi.7")

kpi.t <- function(kpi.sn){
    
    require("dplyr")
    
    # KPI targets
    
    targets <- read.csv(file.path("target", "targets.csv"), stringsAsFactors = FALSE)
    
    # Construct regex from the input kpi. s/n
    
    kpi.regex <- paste("^", kpi.sn, sep = "")
    
    # Filter KPI targets dataframe to return needed targets
    
    target.req <- targets %>% filter(grepl(kpi.regex, KPI)) %>%
                                select(-KPI)
    
    target.req
}


# kpi.1 Access (A&E waiting time) ------------------------------------------------------

kpi.1 <- function(Mmm){
    
kpi_source_helper(Mmm)
    
    path <- grep("(.*kpi.1 .*)", source.kpi$filepath, value = TRUE)
    
    AE_WT <- read_range(path, 5:12, 1:14)
    
    AE_WT.frame <- empty.frame
    
    row.index <- c(2:5)
    
    for (i in 1:length(row.index)){
        AE_WT.frame[i,] <- rep(NA, length(AE_WT.frame))
    }
    
    # Fit processed data to dataframe and show NA data
    
    AE_WT <- AE_WT[row.index,]
    
    # row.names(AE_WT.frame) <- row.names(AE_WT)
    
    for (i in 1:length(AE_WT.frame)){
        if (names(AE_WT.frame)[i] %in% names(AE_WT)){
            AE_WT.frame[i] <- AE_WT[names(AE_WT.frame)[i]]
        }
    }
    
    
    # Replace NA with N.A. for production use in Excel
    
    AE_WT.prod <- as.matrix(AE_WT.frame, rownames.force = TRUE)
    AE_WT.prod <- apply(AE_WT.prod, 2, FUN = function(x){ifelse(is.na(x), "N.A.", x)})
    
    AE_WT.prod <- data.frame(
        lapply(split(AE_WT.prod, col(AE_WT.prod)), type.convert, as.is = TRUE),
        stringsAsFactors = FALSE
    )
    row.names(AE_WT.prod) <- row.names(AE_WT.frame)
    names(AE_WT.prod) <- names(AE_WT.frame)
    
    for (i in 1:length(AE_WT.prod)){
        AE_WT.prod[,i] <- sapply(AE_WT.prod[,i], FUN = function(x) ifelse(is.numeric(x), x/100, x))
    }
    
    # Return production ready dataframe
    
    AE_WT.prod
}

# kpi.2 Access(SOP waiting time) ------------------------------------------

kpi.2 <- function(Mmm, specialty = "Ovr"){
    
    kpi_source_helper(Mmm)
    
#     SOP.specialty <- data.frame(
#         specialty = c("ENT" , "GYN" , "MED" , "OPH" , "ORT" , "PAE" , "PSY" , "SUR"),
#         specialty.df = paste("SOP_WT.", SOP.specialty, sep = ""))
    
    path <- grep("(.*kpi.2 .*)", source.kpi$filepath, value = TRUE)
    path.HA <- grep("(.*kpi.2.HA .*)", source.kpi$filepath, value = TRUE)
    
    SOP_WT <- read_range(path, 5:28, 1:55)
    SOP_WT.HA <- read_range(path.HA, 5:28, 1:10)
        row.index <- c(3:4, 10, 14, 20)

    SOP_WT.frame <- empty.frame
    
            for (i in 1:length(row.index)){
                SOP_WT.frame[i,] <- rep(NA, length(SOP_WT.frame))
            }
    
    # Split data into dataframes per specialty and show NA data
    
    SOP_WT <- SOP_WT[row.index,]
    SOP_WT.HA <- SOP_WT.HA[row.index,]
    
    colnames(SOP_WT) <- gsub("( Mgt)", "", colnames(SOP_WT))
    colnames(SOP_WT.HA) <- gsub("( Inst.$)", "", colnames(SOP_WT.HA))
    
    for (i in 1:(length(SOP_WT)-1)){
        SOP_WT[,i+1] <- as.numeric(SOP_WT[,i+1])
    }
    
    for (i in 1:(length(SOP_WT.HA)-1)){
        SOP_WT.HA[,i+1] <- as.numeric(SOP_WT.HA[,i+1])
    }
    
    ### NEED TO INSERT HA OVERALL DATA ###
    
    if (specialty=="ENT"){
            SOP_WT.Spl <- cbind(SOP_WT[,c(1,2:6)], SOP_WT.HA[,c(2,1)])
    } else if (specialty=="GYN"){
            SOP_WT.Spl <- cbind(SOP_WT[,c(1,7:12)], SOP_WT.HA[,c(3,1)]) 
    } else if (specialty=="MED"){
            SOP_WT.Spl <- cbind(SOP_WT[,c(1,13:19)], SOP_WT.HA[,c(4,1)])
    } else if (specialty=="OPH"){
            SOP_WT.Spl <- cbind(SOP_WT[,c(1,20:23)], SOP_WT.HA[,c(5,1)])
    } else if (specialty=="ORT"){
            SOP_WT.Spl <- cbind(SOP_WT[,c(1,24:30)], SOP_WT.HA[,c(6,1)])
    } else if (specialty=="PAE"){
            SOP_WT.Spl <- cbind(SOP_WT[,c(1,31:46)], SOP_WT.HA[,c(7,1)])
    } else if (specialty=="PSY"){
            SOP_WT.Spl <- cbind(SOP_WT[,c(1,37:40)], SOP_WT.HA[,c(8,1)])
    } else if (specialty=="SUR"){
            SOP_WT.Spl <- cbind(SOP_WT[,c(1,41:47)], SOP_WT.HA[,c(9,1)])
    } else if (specialty=="Ovr"){
            SOP_WT.Spl <- cbind(SOP_WT[,c(1,48:55)], SOP_WT.HA[,c(10,1)])
    }


    for (i in 1:length(SOP_WT.frame)){
        if (names(SOP_WT.frame)[i] %in% names(SOP_WT.Spl)){
            SOP_WT.frame[i] <- SOP_WT.Spl[names(SOP_WT.frame)[i]]
        }
    }
    
    # Convert percentage into decimals
    
    percent_row <- c(3:4)
    
    for (i in 1:length(percent_row)){
            SOP_WT.frame[percent_row[i],] <- sapply(SOP_WT.frame[percent_row[i],], FUN = function(x) ifelse(is.numeric(x), x/100, x))
    }
    
    # Replace NA with N.A. for production use in Excel
    
    SOP_WT.prod <- data.frame(apply(SOP_WT.frame, 2, FUN = function(x){ifelse(is.na(x), "N.A.", x)}), stringsAsFactors = FALSE)
    
    for (i in 1:ncol(SOP_WT.prod)){
        if ("N.A." %in% SOP_WT.prod[,i]){
        } else {
                SOP_WT.prod[,i] <- as.numeric(SOP_WT.prod[,i])
        }
    }
    
    # Return production ready dataframe
    
    SOP_WT.prod
    
}

