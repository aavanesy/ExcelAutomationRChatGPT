library(openxlsx)
library(dplyr)
library(lubridate)

# Fonts 
?openxlsx::modifyBaseFont

# rows and columns
?openxlsx::setColWidths
?openxlsx::setRowHeights

my_headerStyle <- createStyle(
  fontName = "Calibri",
  fontSize = 12,
  fontColour = "black",
  halign = "left",
  valign = 'center',
  textDecoration = 'bold',
  fgFill = "lightgrey",
  border = "TopBottomLeftRight",
  borderColour = "black")

my_bodyStyle <- createStyle(
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

# when reading a csv > mutate date
data <- read.csv('data/Sales_Shops.csv')

str(data)

data <- data %>% 
  mutate(Date = as.Date(Date))

str(data)

data_split <- data %>% 
  dplyr::group_by(Item, SaleType) %>% 
  dplyr::group_split()


date_style = createStyle(numFmt = "mmmm dd, yyyy",
                         valign = 'center',
                         wrapText = T,
                         halign = 'left',
                         textDecoration = 'bold')

# for one sheet
wb <- openxlsx::createWorkbook()

i = 1

data_i <- data_split[[i]]

name_i <- paste(data_i$Item[1], data_i$SaleType[1], sep = ' ')

addWorksheet(wb, sheetName = name_i, gridLines = F)

writeData(wb, sheet = name_i,
          x = data_i, 
          startCol = 2,
          startRow = 4,
          withFilter = T)

addStyle(wb, name_i,
         style = my_bodyStyle,
         rows = 4:(4+nrow(data_i)),
         cols = 2:(1 + ncol(data_i)),
         gridExpand = TRUE)

addStyle(wb, name_i,
         style = my_headerStyle,
         rows = 4, 
         cols = 2:(1 + ncol(data_i)),
         gridExpand = TRUE)


addStyle(wb, name_i, 
         style = createStyle(numFmt = 'DATE'),
         rows = 4:(4+nrow(data_i)), 
         cols = 2,
         gridExpand = TRUE,
         stack = T) #stack function

writeData(wb, sheet = name_i,
          x = paste('Report as of'), 
              startCol = 2,
          startRow = 2)

writeData(wb, sheet = name_i,
          x = Sys.Date(),
          startCol = 3,
          startRow = 2)

addStyle(wb, name_i,
         style = date_style,
         cols = 3,
         rows = 2,
         gridExpand = TRUE, 
         stack = TRUE)

addStyle(wb, name_i,
         style = createStyle(textDecoration = 'bold'),
         cols = 2,
         rows = 2,
         gridExpand = TRUE, 
         stack = TRUE)


mergeCells(wb, name_i, cols = 3:4, rows = 2)

setRowHeights(wb, sheet = name_i, rows = 2, heights = 20)

openXL(wb)


## Multiple pages

wb <- openxlsx::createWorkbook()

for(i in 1:length(data_split)){
  
    print(i)
  
    data_i <- data_split[[i]]
    
    name_i <- paste0(data_i$Item[1], data_i$SaleType[1])
    
    addWorksheet(wb, sheetName = name_i, gridLines = F)
    
    writeData(wb, sheet = name_i,
              x = data_i, 
              startCol = 2,
              startRow = 4,
              withFilter = T)
    
    addStyle(wb, name_i, style = my_bodyStyle,
             rows = 4:(4+nrow(data_i)), cols = 2:(1 + ncol(data_i)),
             gridExpand = TRUE)
    
    addStyle(wb, name_i, style = my_headerStyle,
             rows = 4, cols = 2:(1 + ncol(data_i)),
             gridExpand = TRUE)
    
    
    addStyle(wb, name_i, style = createStyle(numFmt = 'DATE'),
             rows = 4:(4+nrow(data_i)), cols = 2,
             gridExpand = TRUE, stack = T) #stack function
    
    writeData(wb, sheet = name_i,
              x = paste('Report as of'), 
              startCol = 2,
              startRow = 2)
    
    writeData(wb, sheet = name_i,
              x = Sys.Date(),
              startCol = 3,
              startRow = 2)
    
    addStyle(wb, name_i,
             style = createStyle(textDecoration = 'bold', valign = 'center', halign = 'left'),
             cols = 2,
             rows = 2,
             gridExpand = TRUE, 
             stack = TRUE)
    
    addStyle(wb, name_i, 
             style = date_style,
             cols = 3,
             rows = 2,
             gridExpand = TRUE, stack = T)
    
    addStyle(wb, name_i, 
             style = createStyle(borderStyle = 'dashDot', 
                                 borderColour = 'steelblue', 
                                 border =  c("top", "bottom", "left", "right")),
             cols = 2:4,
             rows = 2,
             gridExpand = TRUE, stack = TRUE)
    
    mergeCells(wb, name_i, cols = 3:4, rows = 2)
    
    setRowHeights(wb, sheet = name_i, rows = 2, heights = 30)

}

saveWorkbook(wb, 'output/Report2.xlsx', overwrite = T)




