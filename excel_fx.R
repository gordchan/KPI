
require("dplyr")

hosp <- c("CMC", "KCH", "KWH", "NLTH", "OLMH", "PMH", "WTSH", "YCH", "KWC")
t.fx <- '=IF(G10<>"N.A.", G10-S10, "N.A.")'
fx <- data.frame(hosp, tv = LETTERS[1:9], stringsAsFactors = FALSE)
l <- which(LETTERS=="G")

# Target Variance
for (i in 1:9){
    fx$tv[i] <- gsub(LETTERS[l], LETTERS[l+i], t.fx)
}
# Prior Year Variance
fx <- fx %>% mutate(pyv = gsub("(S)", "W", tv))
# HA Variance
fx <- fx %>% mutate(hav = gsub("(S)", "AA", tv))

write.csv(fx, "excel_fx.csv", row.names = FALSE)
