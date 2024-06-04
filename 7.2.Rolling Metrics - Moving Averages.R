library(openxlsx)
library(dplyr)
library(lubridate)

# when reading a csv > mutate date
data <- read.csv('data/Sales_Shops.csv')

str(data)

data <- data %>% 
  mutate(Date = as.Date(Date))

str(data)

## data summary ----

data <- data %>%
  mutate(Month = format(Date, "%b %Y"))
data$Month <- factor(data$Month, levels = unique(data$Month))

monthly_sale_type <- data %>% 
  group_by(Month, SaleType) %>% 
  summarise(Sales = sum(Price)) 


rolling_sales <- monthly_sale_type %>% 
  arrange(SaleType) %>% 
  group_by(SaleType) %>% 
  mutate(RollingMean3 = zoo::rollapply(Sales, FUN = mean, width = 3, fill = NA, align = 'right')) %>% 
  mutate(RollingMean6 = zoo::rollapply(Sales, FUN = mean, width = 6, fill = NA, align = 'right')) %>% 
  mutate(RollingMean12 = zoo::rollapply(Sales, FUN = mean, width = 12, fill = NA, align = 'right')) %>% 
  ungroup()

library(ggplot2)

rolling_sales %>% 
  tidyr::pivot_longer(-c(1:2)) %>% 
  ggplot(aes(x = Month, y = value, color = name, group = 1)) + 
  geom_point() + 
  facet_wrap(~SaleType, nrow = 3)

## Analytics ----
wb <- openxlsx::createWorkbook()
sheet_i = 'Analytics'
addWorksheet(wb, sheetName = sheet_i, tabColour = 'steelblue')

writeData(wb, sheet = sheet_i,
          x = rolling_sales, 
          # headerStyle = my_headerStyle,
          borders = c('all'),
          borderColour = 'grey',
          borderStyle = 'thin',
          withFilter = TRUE)


addStyle(wb, sheet_i, 
         style = createStyle(numFmt = '#,##0.0'),
         rows = 2:(1+nrow(rolling_sales)), 
         cols = 3:ncol(rolling_sales),
         gridExpand = TRUE, stack = T)

rows_grey <- rolling_sales %>% 
  ungroup() %>% 
  mutate(rid = 1:n()) %>% 
  mutate(mm = grepl('Mar|Jun|Sep|Dec', Month)) %>% 
  filter(mm) %>% 
  pull(rid) + 1
  
addStyle(wb, sheet_i, 
         style = createStyle(fgFill = "#4F81BD"),
         rows = rows_grey, 
         cols = 1:ncol(rolling_sales),
         gridExpand = TRUE, stack = T)


openxlsx::openXL(wb)




# rolling_sales %>% 
#   split(f = as.factor(.$SaleType)) %>% 
#   openxlsx::write.xlsx(file = 'output/RollingReturns.xlsx',
#                        asTable = T,
#                        tableStyle = 'TableStyleLight17',
#                        overwrite = TRUE)
