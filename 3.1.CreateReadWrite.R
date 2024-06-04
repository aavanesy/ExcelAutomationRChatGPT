
library(openxlsx)


# create a workbook and save it ----
?createWorkbook
wb <- openxlsx::createWorkbook()
#wb <- openxlsx::createWorkbook(creator = 'Arkadi')
print(wb)

?saveWorkbook
openxlsx::saveWorkbook(wb, file = 'data/file.xlsx', overwrite = TRUE, returnValue = TRUE)


# read and save it ----
?loadWorkbook
wb <- loadWorkbook(file = 'data/Sales_Shops_formatted.xlsx')
print(wb)
class(wb)
## ... make changes...

saveWorkbook(wb, 'data/Sales_Shops2.xlsx', overwrite = T, returnValue = TRUE)


# add sheets and data, view before saving ----

wb <- openxlsx::createWorkbook()

?addWorksheet
addWorksheet(wb, sheetName = 'Summary', tabColour = 'grey')

openxlsx::openXL(wb)

?writeData
writeData(wb, sheet = 'Summary', x = iris[1:10,])

openxlsx::openXL(wb)

openxlsx::saveWorkbook(wb, 'data/file.xlsx', overwrite = T)
