# conditional formatting ----

?openxlsx::conditionalFormatting

data <- read.csv('data/Sales_Shops.csv') %>% 
  tail(50)

data

wb <- openxlsx::createWorkbook()

wsheet = 'Sales'

addWorksheet(wb, sheetName = wsheet)

writeData(wb, wsheet, data)

openxlsx::openXL(wb)

# Conditional Formatting --------

?conditionalFormatting


avg_val = mean(data$Price)
rule_v = paste0(">=",avg_val)

conditionalFormatting(
  wb,
  wsheet,
  cols = 6,
  rows = 2:(nrow(data)+1),
  rule = rule_v,
  style = createStyle(fontColour = "#FC0000", bgFill = "#C6EFCE"),
  type = "expression"
)
openxlsx::openXL(wb)


conditionalFormatting(
  wb,
  wsheet,
  cols = 6,
  rows = 2:(nrow(data)+1),
  rule = NULL,
  style = "#31869B",
  type = "databar"
)
openxlsx::openXL(wb)

conditionalFormatting(
  wb,
  wsheet,
  cols = 6,
  rows = 2:(nrow(data)+1),
  rule = NULL,
  style = "#31869B",
  type = "databar"
)

##? how to remove?




