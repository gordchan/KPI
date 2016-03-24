# HTML data direct
#
# To load system generated html page disguised as .xls file
#
# March 2016

# read_Chtml -----------------------------------------------------------------------

read_Chtml <- function (input, hi){
    
    require(htmltab)
    require(XML)
    
# Acquire html lines
    
    con <- file(input, "r")
    lines <- readLines(con)
    close(con)
    rm(input)
    rm(con)
    
# File formatting
    
    # Indexing
    A <- 1:4 # Header
    B <- 5 # Name
    E <- (length(lines)-6):length(lines) # Disclaimer
    
    D <- which(grepl("<table id", lines))[length(which(grepl("<table id", lines)))]:(E[1]-1) # Cream of the Crops Datatable (if multiple datatable the last will be used)
    C <- 6:(D[1]-1) # Quere details
    
    lines <- lines[c(B, D)]
    
    lines <- gsub("<\\/?font.{0,20}>", "", lines)
    lines <- gsub("<tr", "<\\/tr><tr", lines)
    lines[2] <- sub("<\\/tr>", "", lines[2])
    
# Parse and return dataframe
    
    hi <<- hi

    parse <- htmlParse(lines, asText = TRUE)
    df <- htmltab(parse, which = "//table[@id='datatable']", header = hi)
    
# Return df
    
    df
}