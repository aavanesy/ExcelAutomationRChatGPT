
library(openxlsx)
library(dplyr)

data <- read.csv('data/Sales_Shops.csv') %>% 
  slice_head(n = 20)

## create custom styles

?createStyle

?addStyle

headerStyleA <- createStyle(
  fontName = "Calibri",
  fontSize = 12,
  fontColour = "black",
  halign = "left",
  valign = 'center',
  textDecoration = 'bold',
  fgFill = "lightgrey",
  border = "TopBottomLeftRight",
  borderColour = "black"
)

class(headerStyle)
print(headerStyle)

## workbook code ----
wb <- openxlsx::createWorkbook()

sheet_i <- 'Data'

addWorksheet(wb, sheetName = sheet_i, gridLines = FALSE)

writeData(wb, sheet = sheet_i, x = data)
# writeData(wb, sheet = sheet_i, x = data, headerStyle = headerStyleA)

openxlsx::openXL(wb)

addStyle(wb, sheet_i, rows = 1, cols = 1:ncol(data), style = headerStyle)

openxlsx::openXL(wb)

bodyStyle <- createStyle(
  fontName = "Calibri",
  fontSize = 11,
  fontColour = 'black',
  numFmt = "GENERAL",
  border = 'TopBottomLeftRight',
  borderColour = "darkgrey",
  borderStyle = "thin" ,
  bgFill = NULL,
  fgFill = '#FBFBFB',
  halign = 'left',
  valign = 'center')

addStyle(wb, sheet_i, style = bodyStyle,
         rows = 2:(1+nrow(data)), cols = 1:ncol(data),
         gridExpand = TRUE)

openxlsx::openXL(wb)



## ChatGPT example

# Load the openxlsx package

# Sample data frame
data <- data.frame(
  Name = c("Alice", "Bob", "Charlie", "David"),
  Age = c(23, 35, 45, 28),
  Score = c(89, 92, 85, 88)
)

# Create a workbook and add a worksheet
wb <- createWorkbook()
addWorksheet(wb, "Sheet 1")

# Write the data to the worksheet
writeData(wb, sheet = "Sheet 1", x = data, startRow = 1, startCol = 1)

# Define the styles
headerStyle <- createStyle(fontColour = "#FFFFFF", fgFill = "#4F81BD", halign = "CENTER", textDecoration = "BOLD")
dataStyle <- createStyle(fgFill = "#DCE6F1")

# Apply the header style
addStyle(wb, sheet = "Sheet 1", style = headerStyle, rows = 1, cols = 1:ncol(data), gridExpand = TRUE)

# Apply the data style (excluding the header)
addStyle(wb, sheet = "Sheet 1", style = dataStyle, rows = 2:(nrow(data) + 1), cols = 1:ncol(data), gridExpand = TRUE)

# Save the workbook
saveWorkbook(wb, file = "data/styled_table.xlsx", overwrite = TRUE)
