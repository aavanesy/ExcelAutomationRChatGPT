

setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

# Scrap data from https://www.slickcharts.com/dowjones
url <- 'https://www.slickcharts.com/dowjones'

# libraries
library(rvest)
library(dplyr)
library(httr)

# Read the HTML content from the URL
page <- read_html(url)

# Extract the table from the HTML content
table <- page %>%
  html_node("table") %>%
  html_table(fill = TRUE)

# convert last column to numeric
table <- table %>% 
  mutate(`% Chg` = gsub('\\(', '', `% Chg`)) %>% 
  mutate(`% Chg` = gsub('\\)', '', `% Chg`)) %>% 
  mutate(`% Chg` = gsub('\\%', '', `% Chg`)) %>% 
  mutate(`% Chg` = as.numeric(`% Chg`) / 100) %>% 
  select(-1) %>% 
  mutate(Price = ifelse(is.nan(Price), NA, Price)) %>% 
  mutate(`% Chg` = ifelse(is.nan(`% Chg`), NA, `% Chg`))

# numeric conversions
table <- table %>% 
  mutate(Weight = Weight / 100)

class(table$Weight) = 'percentage'
class(table$`% Chg`) = 'percentage'
class(table$Price) = 'numeric'
class(table$Chg) = 'numeric'
str(table)

# write Excel Report
library(openxlsx)

# Simple version ----
write.xlsx(table, file = 'DowJonesReport.xlsx',
           asTable = TRUE, tableStyle = 'TableStyleMedium6')


# Advanced version ----
file_name <- Sys.Date() %>% format('%d.%m.%Y')
file_name <- paste0('DowJonesReport_',file_name)
file_name <- paste0(file_name, '.xlsx')

wb <- createWorkbook()
addWorksheet(wb, 'Report', gridLines = FALSE)

openxlsx::modifyBaseFont(wb, fontSize = 10, fontName = 'Calibri')

my_headerStyle <- createStyle(
  fontSize = 12,
  fontColour = "black",
  halign = "left",
  valign = 'center',
  textDecoration = 'bold',
  fgFill = "#EEECE1",
  border = "TopBottomLeftRight",
  borderColour = "black")

my_bodyStyle <- createStyle(
  fontColour = '#26292C',
  border = 'TopBottomLeftRight',
  borderColour = "darkgrey",
  borderStyle = "thin" ,
  fgFill = '#FBFBFB',
  halign = 'left',
  valign = 'center')

writeData(wb, 'Report', x = table, headerStyle = my_headerStyle)

addStyle(wb, 'Report', 
         rows = 2:(1+nrow(table)),
         cols = 1:ncol(table),
         style = my_bodyStyle,
         gridExpand = TRUE,
         stack = TRUE)

addStyle(wb, 'Report', 
         rows = 2:(1+nrow(table)),
         cols = c(3,6),
         style = createStyle(numFmt = 'PERCENTAGE'),
         gridExpand = TRUE,
         stack = TRUE)

addStyle(wb, 'Report', 
         rows = 2:(1+nrow(table)),
         cols = 4:5,
         style = createStyle(numFmt = 'NUMBER'),
         gridExpand = TRUE,
         stack = TRUE)

setColWidths(wb, 'Report', cols = 1, widths = 30)
setRowHeights(wb, 'Report', rows = 1, heights = 18)

openXL(wb)

saveWorkbook(wb, file = file_name)



