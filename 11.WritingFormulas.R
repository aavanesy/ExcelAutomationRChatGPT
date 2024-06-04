library(openxlsx)
library(dplyr)
library(lubridate)

setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

# when reading a csv > mutate date
data <- read.csv('data/Sales_Shops.csv') %>% head(200)

str(data)

data <- data %>% 
  mutate(Date = as.Date(Date))

?writeFormula

wb <- openxlsx::createWorkbook()
addWorksheet(wb, 'RawData')
writeData(wb, 'RawData', data)

result <- data %>%
  group_by(Country) %>%
  summarise(TotalPrice = sum(Price))

writeData(wb, 'RawData', x = result, startCol = 8, startRow = 3)

openXL(wb)

# formulas <- c('SUMIF(D2:D201,H4,F2:F201)') 

f1 = paste0('SUMIF(D2:D', nrow(data) + 1,
            ',H4,F2:F',
            nrow(data) + 1,
            ')')
formulas <- c('SUMIF(D2:D201,H4,F2:F201)') 

## LOOP VERSION
# formulas = c()
# for(i in 1:length(unique(data$Country))){
#   print(i)
#   f1 = paste0('SUMIF(D2:D', nrow(data) + 1,
#               ',H',
#               i + 3,
#               ',F2:F',
#               nrow(data) + 1,
#               ')')
#   print(f1)
#   formulas = c(formulas, f1)
# }


writeFormula(wb, 'RawData', x = formulas, startCol = 10, startRow = 4)

openXL(wb)


library(glue)
original_formula <- "=SUMIF(D2:D201, G{cell}, F2:F201)"
cell = 50
glue("=SUMIF(D2:D201, G{cell}, F2:F201)")


?makeHyperlinkString

## Writing internal hyperlinks
wb <- createWorkbook()
addWorksheet(wb, "Sheet1")
addWorksheet(wb, "Sheet2")
addWorksheet(wb, "Sheet 3")
writeData(wb, sheet = 3, x = iris)

## External Hyperlink
x <- c("https://www.google.com", "https://www.google.com.au")
names(x) <- c("google", "google Aus")
class(x) <- "hyperlink"

writeData(wb, sheet = 1, x = x, startCol = 10)
openXL(wb)
