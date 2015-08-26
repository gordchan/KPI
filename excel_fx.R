
require("dplyr")

hosp <- c("CMC", "KCH", "KWH", "NLTH", "OLMH", "PMH", "WTSH", "YCH", "KWC")


# S, W ----------------------------------------------------------------

l <- which(LETTERS=="R")
L <- LETTERS[l]
Ln <- LETTERS[l+9]

t.fx <- '=source!'
fx <- data.frame(t = LETTERS[1:25], stringsAsFactors = FALSE)

for (i in 1:25){
        fx$t[i] <- paste(t.fx, LETTERS[l], 3+i, sep = "")
}

fx <- fx %>% mutate(py = gsub(LETTERS[l], LETTERS[l+9], t))

write.csv(fx$t, paste("SW/", L, ".csv", sep = ""), row.names = FALSE, quote = FALSE)
write.csv(fx$py, paste("SW/", Ln, ".csv", sep = ""), row.names = FALSE, quote = FALSE)

#  T, X, AB ---------------------------------------------------------------

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
