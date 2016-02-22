# Duplicate queues
#
# Copy source queues containing 24-months data

dup_queue <- function(y, m){

# Fx announcement ------------------------------------

print(paste("Duplicating Queues Started:", y, "-", m, Sys.time()))

# Libraries ------------------------------------------

source("scripts/kpi_dates.R")
require(dplyr)

# Dates ----------------------------------------------

KPI_dates(y, m)

# MIPO
    
    ls_MIPO <- list.files(file.path("source", Dates[1,], "MIPO"))
    ls_MIPO <- grep("kpi.6", ls_MIPO, value = TRUE)

    for (i in 1:length(ls_MIPO)){
        file.copy(from = file.path("source", Dates[1,], "MIPO", ls_MIPO[i]),
        to = file.path("source", Dates[3,], "MIPO", ls_MIPO[i]),
        overwrite = TRUE)
        
    }
    
# RAD
    
    dir.create(path = file.path("source", Dates[3,], "RAD"),
               showWarnings = FALSE, recursive = TRUE)
    
    file.copy(from = file.path("source", Dates[1,], "RAD"),
              to = file.path("source", Dates[3,]),
              overwrite = TRUE, recursive = TRUE)
    
# Fx end announcenment ---------------------------
    
    print(paste("Duplicating Queues End:", y, "-", m, Sys.time()))    

}




# Standardise Filenames
#
# Rename source queues into stdand format and copy to root dir of the period.


std_filename <- function(y, m){

# Fx announcement ------------------------------------
    
    print(paste("Filename Standardisation Started:", y, "-", m, Sys.time()))
    
# Libraries ------------------------------------------
    
    source("scripts/kpi_dates.R")
    require(dplyr)

# Dates ----------------------------------------------

KPI_dates(y, m)

# Folder paths ---------------------------------------

f_CDARS <- file.path("source", Dates[1,], "CDARS")
f_EIS <- file.path("source", Dates[1,], "EIS")
f_MIPO <- file.path("source", Dates[1,], "MIPO")
f_RAD <- file.path("source", Dates[1,], "RAD")


# CDARS queues ---------------------------------------

fs_CDARS <- data.frame(filename_ori = list.files(f_CDARS),
                     filename_to = NA,
                     stringsAsFactors = FALSE)
fs_CDARS <- fs_CDARS %>% filter(grepl("xlsx$", filename_ori))

fs_CDARS$filename_to <- fs_CDARS$filename_ori # Get kpi/tre numbers
        fs_CDARS$filename_to <- sub(".*_kpi([1-9]+).*", "kpi.\\1", fs_CDARS$filename_to)
        fs_CDARS$filename_to <- sub(".*_tre([1-9]+).*", "tre.\\1", fs_CDARS$filename_to)
        fs_CDARS$filename_to <- sub(".*_mrsa.*", "kpi.3.3", fs_CDARS$filename_to)
    
        fs_CDARS$filename_to <- gsub("(.*)([1-9])([1-9])([1-9])", "\\1\\2\\.\\3\\.\\4", fs_CDARS$filename_to) # 3 digits
        fs_CDARS$filename_to <- gsub("(.*)([1-9])([1-9])", "\\1\\2\\.\\3", fs_CDARS$filename_to) # 2 digit

for (i in 1:nrow(fs_CDARS)){
    fs_CDARS$filename_to[i] <- paste(fs_CDARS$filename_to[i], fs_CDARS$filename_ori[i])
}
        
        
for (i in 1:nrow(fs_CDARS)){
    
    file.copy(from = file.path("source", Dates[1,], "CDARS", fs_CDARS$filename_ori[i]),
              to = file.path("source", Dates[1,], fs_CDARS$filename_to[i]),
              overwrite = TRUE)
    
} 
        
# EIS queues -----------------------------------------

fs_EIS <- list.files(f_EIS)

for (i in 1:length(fs_EIS)){
    
    file.copy(from = file.path("source", Dates[1,], "EIS", fs_EIS[i]),
              to = file.path("source", Dates[1,], fs_EIS[i]),
              overwrite = TRUE)
    
}

# MIPO queues ---------------------------------------

fs_MIPO <- list.files(f_MIPO)

for (i in 1:length(fs_MIPO)){
    
    file.copy(from = file.path("source", Dates[1,], "MIPO", fs_MIPO[i]),
              to = file.path("source", Dates[1,], fs_MIPO[i]),
              overwrite = TRUE)
    
}

# RAD queues ---------------------------------------

fs_RAD <- data.frame(filename_ori = list.files(f_RAD),
                     filename_to = NA,
                     stringsAsFactors = FALSE)

# kpi.7 Radiology Waiting Time

    x <- grep("KWC_Radiology_Waiting_time", fs_RAD$filename_ori)
    
    fs_RAD$filename_to[x] <- paste("kpi.7 ", fs_RAD$filename_ori[x])
    
# tre.12 Trend of Retrocpective Radiology Waiting Time
    
    x <- grep("KWC_Trend", fs_RAD$filename_ori)
    
    fs_RAD$filename_to[x] <- paste("tre.12 ", fs_RAD$filename_ori[x])
    
# Fileter NA
    
    fs_RAD <- fs_RAD %>% filter(!is.na(filename_to))

    
for (i in 1:length(fs_RAD)){
    
    file.copy(from = file.path("source", Dates[1,], "RAD", fs_RAD$filename_ori[i]),
              to = file.path("source", Dates[1,], fs_RAD$filename_to[i]),
              overwrite = TRUE)
    
}
    
# Fx end announcenment ---------------------------
    
    print(paste("Filename Standardisation End:", y, "-", m, Sys.time()))
    
}