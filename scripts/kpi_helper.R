
# Helper function to help with reading source file ----------------------------------------

# i. Initialise empty dataframe for use by each measurements
#
# ii. Generate file table according to reporting month period specified

kpi_source_helper <- function(Mmm){
    
    require("xlsx")
    require("dplyr")
    
    if(!exists("read_range", mode="function")) source("scripts/read_xlsx.R")
    if(!exists("read_html", mode="function")) source("scripts/read_html.R")
    
    # Empty dataframe
    
    empty.frame <<- data.frame(CMC = numeric(0),
                               KCH = numeric(0),
                               KWH = numeric(0),
                               NLTH = numeric(0),
                               OLM = numeric(0),
                               PMH = numeric(0),
                               WTS = numeric(0),
                               YCH = numeric(0),
                               KWC = numeric(0),
                               HA = numeric(0))
    
    # File table
    # KPI targets
    
    target.kpi <<- data.frame(filename = c("targets.csv")) %>%
        mutate(filepath = file.path("target", Mmm, filename))
    
    
    # Clinical KPI
    
#     source.HA.kpi <- read.xlsx("source/KPI items.xlsx",
#                                sheetName = "KPI",
#                                colIndex = 5,
#                                stringsAsFactors = FALSE)
#     names(source.HA.kpi) <- "Query.Name"
    
    source.kpi <<- data.frame(filename = list.files(file.path("source", Mmm), pattern = ".*\\.(xls.?|html)")) %>%
        mutate(filepath = file.path("source", Mmm, filename)) ## USED BY OTHER fx
    
    # KPI Trend
#     source.tre <- read.xlsx("source/KPI items.xlsx",
#                             sheetName = "KPI",
#                             colIndex = 4,
#                             stringsAsFactors = FALSE)
#     
#     source.HA.tre <- read.xlsx("source/KPI items.xlsx",
#                                sheetName = "KPI",
#                                colIndex = 5,
#                                stringsAsFactors = FALSE)
#     names(source.HA.tre) <- "Query.Name"
    
#     source.tre <- bind_rows(source.tre, source.HA.tre)
#     
#     source.tre <<- source.tre %>% filter(!is.na(Query.Name)) %>%
#         distinct(Query.Name) %>% arrange(Query.Name) %>%
#         mutate(filename = paste(Query.Name, ".xlsx", sep = "")) %>%
#         mutate(filepath = file.path("source", Mmm, filename))    
    
    # SOP KPI
#     source.sop <- read.xlsx("source/KPI items.xlsx",
#                             sheetName = "KPI",
#                             colIndex = 4,
#                             stringsAsFactors = FALSE)
#     
#     source.HA.sop <- read.xlsx("source/KPI items.xlsx",
#                                sheetName = "KPI",
#                                colIndex = 5,
#                                stringsAsFactors = FALSE)
#     names(source.HA.sop) <- "Query.Name"
#     
#     source.sop <- bind_rows(source.sop, source.HA.sop)
#     
#     source.sop <<- source.sop %>% filter(!is.na(Query.Name)) %>%
#         distinct(Query.Name) %>% arrange(Query.Name) %>%
#         mutate(filename = paste(Query.Name, ".xlsx", sep = "")) %>%
#         mutate(filepath = file.path("source", Mmm, filename))    
}