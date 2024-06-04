library(dplyr)
library(lubridate)
library(openxlsx)

start_date = as.Date('2020-01-01')
end_date = as.Date('2023-01-31')

# nsales = 10

df <- data.frame(s1 = 1:10 / 7,
                 s2 = 1:10 / 7,
                 s3 = 1:10 / 7,
                 s4 = 1:10 / 7,
                 s5 = seq(1000, 1009, by  = 1),
                 s6 = 1:10 / 7,
                 s7 = 1:10 / 7)

## Excel code ----

wb <- createWorkbook()
sheet_x = 'Numbers'
addWorksheet(wb, sheet_x)
writeData(wb, sheet_x, df)

openXL(wb)

# col one
addStyle(wb, sheet_x,
         style = createStyle(numFmt = "GENERAL"),
         rows = 2:11, 
         cols = 1, 
         gridExpand = TRUE)

addStyle(wb, sheet_x,
         style = createStyle(numFmt = "NUMBER"),
         rows = 2:11, 
         cols = 2, 
         gridExpand = TRUE)

openXL(wb)


s <- createStyle(numFmt = "0.00")
addStyle(wb, 1, style = s, rows = 2:6, cols = 3, gridExpand = TRUE)

s <- createStyle(numFmt = "0.000")
addStyle(wb, 1, style = s, rows = 2:6, cols = 4, gridExpand = TRUE)

s <- createStyle(numFmt = "#,##0")
addStyle(wb, 1, style = s, rows = 2:6, cols = 5, gridExpand = TRUE)

s <- createStyle(numFmt = "#,##0.00")
addStyle(wb, 1, style = s, rows = 2:6, cols = 6, gridExpand = TRUE)

s <- createStyle(numFmt = "$ #,##0.00")
addStyle(wb, 1, style = s, rows = 2:6, cols = 7, gridExpand = TRUE)

openXL(wb)
