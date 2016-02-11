KPI_dates <- function(y, m){

    require(lubridate)
    require(dplyr)
    
    to.YY <- paste(unlist(strsplit(as.character(y), ""))[3:4], collapse = "") ## Input from func
    from.YY <- paste(unlist(strsplit(as.character(y-1), ""))[3:4], collapse = "") ## Input from func
    prev.from.YY <- paste(unlist(strsplit(as.character(y-2), ""))[3:4], collapse = "") ## Input from func
    
    to.month <- m ## Input from func
    to.Mmm <- month(to.month, label = TRUE, abbr = TRUE)
    from.Mmm <- month(to.month + 1, label = TRUE, abbr = TRUE)
    
    to.MmmYY <- paste(to.Mmm , to.YY, sep = "")
    to.MmmYY_ <- paste(" ", to.MmmYY, sep = "") # Keep leading zero for use as text in Excel
    
    from.MmmYY <- paste(from.Mmm, from.YY, sep = "")
    
    prev.to.MmmYY <- paste(to.Mmm, from.YY, sep = "")
    prev.from.MmmYY <- paste(from.Mmm, prev.from.YY, sep = "")
    
    cy.period <- paste(from.MmmYY, to.MmmYY, sep = "-")
    py.period <- paste(prev.from.MmmYY, prev.to.MmmYY, sep = "-")
    
    eom <- ymd(paste(y, "-", m , "-", 1)) # Find end date of reporting period this year
    month(eom) <- month(eom) + 1
    day(eom) <- day(eom) -1
    
    som <- eom # Find start date of reporting period this year
    year(som) <- year(som) -1
    day(som) <- day(som) +1
    
    
    date.period <- paste(day(som), month(som, label = TRUE, abbr = TRUE), year(som), "-", day(eom), month(eom, label = TRUE, abbr = TRUE), year(eom))
    date.eom <- paste(day(eom), "/", month(eom), "/", year(eom), sep = "")
    
    date.eom.prev <- paste(day(eom), "/", month(eom), "/", year(eom)-1, sep = "")
    
    df <- data.frame(dates = c(to.MmmYY, # 1 Reporting month
                               to.MmmYY_, # 2 Reporting month w/ leading zero
                               prev.to.MmmYY, # 3 Last year's reporting month
                               cy.period, # 4 Current year reporting period
                               py.period,# 5 Previous year reporting period
                               date.period, # 6 Reporting period in dates (for Excel report)
                               date.eom, # 7 EOM date (for Excel report)
                               date.eom.prev, # 8 EOM date (for Excel report)
                               y, # 9 Raw Reporting Year
                               m # 10 Raw Reporting Month
    ), 
    stringsAsFactors = FALSE)
    
    Dates <<- df

## Data Series for trend

    
    eos <- ymd(paste(y, sprintf("%02.f", m), "01", sep = ""))
    
    series <- data.frame(date = .POSIXct(rep(NA, 12)))
    
    for(i in 0:11){
        series[12-i, 1] <- eos - months(i)
    }
    
    series <- series %>% mutate(month = sapply(date, function(x) month(x, label = TRUE))) %>%
        mutate(year = sapply(date, function(x) paste(unlist(strsplit(as.character(year(x)), ""))[3:4], collapse = ""))) %>%
        mutate(label = paste(month, year))
    
    seriesLabel <- series %>% select(label) %>% t()
    
    Series <<- seriesLabel

}