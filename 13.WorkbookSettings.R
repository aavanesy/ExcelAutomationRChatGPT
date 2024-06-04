
library(openxlsx)

?openxlsx::sheetVisibility

wb <- createWorkbook()
addWorksheet(wb, sheetName = "S1", visible = FALSE)
addWorksheet(wb, sheetName = "S2", visible = TRUE)
addWorksheet(wb, sheetName = "S3", visible = FALSE)

sheetVisibility(wb)

sheetVisibility(wb)[1] <- TRUE ## show sheet 1
sheetVisibility(wb)[2] <- FALSE ## hide sheet 2
sheetVisibility(wb)[3] <- "hidden" ## hide sheet 3
sheetVisibility(wb)[3] <- "veryHidden" ## hide sheet 3 from UI

sheetVisibility(wb)
openXL(wb)
# ?openxlsx::sheetVisible

?openxlsx::protectWorkbook

wb <- createWorkbook()
addWorksheet(wb, "S1")
protectWorkbook(wb, protect = TRUE, password = "Password", lockStructure = TRUE)
openXL(wb)

?openxlsx::protectWorksheet

?openxlsx::worksheetOrder

wb <- createWorkbook()
addWorksheet(wb, sheetName = "S1", visible = TRUE)
addWorksheet(wb, sheetName = "S2", visible = TRUE)
addWorksheet(wb, sheetName = "S3", visible = TRUE)
worksheetOrder(wb)
worksheetOrder(wb) = c(3,1,2)
openXL(wb)

?openxlsx::addCreator
wb <- createWorkbook()
addCreator(wb, "Arkadi")
openxlsx::getCreators(wb)
saveWorkbook(wb, file = 'data/creator.xlsx')

?openxlsx::setLastModifiedBy


