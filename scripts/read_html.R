# HTML data direct
#
# To load html page disguised as .xls file
#
# Dec 2015


# read_html -----------------------------------------------------------------------

read_html <- function (input){
    
    require("XML")
    require("dplyr")
    
    # Read spreadsheet with deliberate over read to accomodate variations in table length
    
    html_data <- readHTMLTable(source.xls, header = FALSE)
    
    html_data <- html_data$datatable
    
    html_data
}