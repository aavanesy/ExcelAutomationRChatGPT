
library(openxlsx)
library(dplyr)

# write table with basic styling ----

# control shift M or command shift M
data <- read.csv('data/Sales_Shops.csv') %>% 
  slice_head(n = 20) %>% 
  distinct()

### Code -----
wb <- openxlsx::createWorkbook()

sheet_i <- 'Data'

addWorksheet(wb, sheetName = sheet_i, gridLines = FALSE, tabColour = 'grey')

openxlsx::openXL(wb)

writeData(wb, sheet = sheet_i, x = data)

openxlsx::openXL(wb)

writeData(wb, sheet = sheet_i,
          x = data,
          startCol = 1,
          startRow = 1,
          borders = 'surrounding',
          borderColour = 'grey',
          borderStyle = 'thin')

openxlsx::openXL(wb)

writeData(wb, sheet = sheet_i,
          x = data,
          startCol = 1,
          startRow = 1,
          borders = c('all'),
          borderColour = 'grey',
          borderStyle = 'thin',
          withFilter = TRUE)

openxlsx::openXL(wb)


