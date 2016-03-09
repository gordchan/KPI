# HTML data direct
#
# To load html page disguised as .xls file
#
# Dec 2015


# read_html -----------------------------------------------------------------------

read_html <- function (input){
    
    require(htmltab)
    
    # Read & parse HTML as df
    
    # html <- file.path("Test", "simple.html")
    
    parse <- htmlParse(input, validate = TRUE)
    
    df <- htmltab(parse, which = "//table[@id='datatable']", header = 1, complementary = FALSE)
    
    names(df) <- c(1:ncol(df))
    
    # tidy corrupted last column, truncate repetition of next row data
    
    for (i in 1:nrow(df)){
        # regex: ([[:alnum:]| ]*?) *KEY.*
        
        key <- df[i+1,1]
        rx <- paste("([[:alnum:]| ]*?) *", key, ".*", sep = "")
        
        df[i,ncol(df)] <- gsub(rx, "\\1", df[i,ncol(df)])
        
    }
    
    # Return df
    
    df
}