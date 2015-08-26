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

kpi_source_helper <- function(m){
    
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
    
    source.kpi <<- read.xlsx("source/KPI items.xlsx",
                            sheetName = "KPI",
                            colIndex = 4)
        source.kpi <<- source.kpi %>% filter(!is.na(Query.Name)) %>%
            mutate(filename = paste(Query.Name, ".xlsx", sep = "")) %>%
            mutate(filepath = file.path("source", m, filename))
    
    source.tre <<- read.xlsx("source/KPI items.xlsx",
                            sheetName = "TRE",
                            colIndex = 4)
        source.tre <<- source.tre %>% filter(!is.na(Query.Name)) %>%
            mutate(filename = paste(Query.Name, ".xlsx", sep = "")) %>%
            mutate(filepath = file.path("source", m, filename))    
    
    source.sop <<- read.xlsx("source/KPI items.xlsx",
                            sheetName = "SOP",
                            colIndex = 4)
        source.sop <<- source.sop %>% filter(!is.na(Query.Name)) %>%
            mutate(filename = paste(Query.Name, ".xlsx", sep = "")) %>%
            mutate(filepath = file.path("source", m, filename))    
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

kpi.1 <- function(m = report.month){
    
kpi_source_helper(m)
    
    path <- grep("(.*kpi.1.*)", source.kpi$filepath, value = TRUE)
    
    AE_WT <- read_range(path, 5:12, 1:14)
    
    AE_WT.frame <- empty.frame
    
    row.index <- c(2:5)
    
    for (i in 1:length(row.index)){
        AE_WT.frame[i,] <- rep(NA, length(AE_WT.frame))
    }
    
    # Fit processed data to dataframe and show NA data
    
    AE_WT <- AE_WT[row.index,]
    
    row.names(AE_WT.frame) <- row.names(AE_WT)
    
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
