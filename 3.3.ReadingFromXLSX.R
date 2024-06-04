
library(openxlsx)
library(dplyr)

setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

?read.xlsx

data_x <- read.xlsx('data/MultipleSheetsTables.xlsx', sheet = 1)

?loadWorkbook

wb <- loadWorkbook('data/MultipleSheetsTables.xlsx')
?names
names(wb)

## Sheet 1 ----

df_one <- read.xlsx('data/MultipleSheetsTables.xlsx', sheet = 1)
df_one <-  read.xlsx(wb, sheet = 1)

read.xlsx('data/MultipleSheetsTables.xlsx',
          sheet = 'Sheet 1',
          # startRow = 4,
          colNames = T,
          sep.names = '_',
          detectDates = T,
          check.names = T)

?openxlsx::convertToDate


# Sheet two ----

read.xlsx(wb, sheet = 'Sheet2')

wb %>% read.xlsx(sheet = 'Sheet2')

df <- read.xlsx(wb, sheet = 'Sheet2', skipEmptyCols = FALSE)


df %>% 
  janitor::remove_empty(which = 'cols')

df <- read.xlsx(wb,
                sheet = 'Sheet2',
                detectDates = TRUE,
                cols = 1:6,
                skipEmptyCols = FALSE)

str(df)

## Sheet Three ----

read.xlsx(wb, sheet = 'Sheet3')

read.xlsx(wb, sheet = 'Sheet3', skipEmptyRows = FALSE)

read.xlsx(wb, sheet = 'Sheet3',
          rows = 1:20,
          skipEmptyRows = FALSE)

