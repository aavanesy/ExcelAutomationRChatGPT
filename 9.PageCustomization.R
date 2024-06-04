# non covered important functions ----

library(openxlsx)

data <- read.csv('data/Sales_Shops.csv') %>% head(16)

wb <- createWorkbook()
addWorksheet(wb, 'Report')
writeData(wb, "Report", data, withFilter = TRUE)

openXL(wb)

## Freezing Panes ----

?openxlsx::freezePane

freezePane(
  wb,
  sheet = 'Report',
  # firstActiveRow = NULL,
  # firstActiveCol = NULL,
  firstRow = TRUE,
  firstCol = TRUE
)

#freezePane(wb, 4, firstActiveRow = 1, firstActiveCol = "D")

openXL(wb)

# Page Setup ----

wb <- createWorkbook()
addWorksheet(wb, 'Report')
writeData(wb, "Report", head(data, 10), withFilter = TRUE)

?openxlsx::pageSetup

pageSetup(
  wb,
  sheet = 'Report',
  orientation = 'landscape',
  scale = 150,
  left = 0.7,
  right = 0.7,
  top = 0.75,
  bottom = 0.75,
  header = 0.3,
  footer = 0.3,
  fitToWidth = FALSE,
  fitToHeight = FALSE,
  paperSize = NULL,
  printTitleRows = NULL,
  printTitleCols = NULL,
  summaryRow = NULL,
  summaryCol = NULL
)

openXL(wb)




