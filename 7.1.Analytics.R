library(openxlsx)
library(dplyr)
library(lubridate)

# when reading a csv > mutate date
data <- read.csv('data/Sales_Shops.csv')

str(data)

data <- data %>% 
  mutate(Date = as.Date(Date))

str(data)


# Main Metrics - Aggregations - Sum - Mean - Median - Percentiles

## data summary ----

data <- data %>%
  mutate(Month = format(Date, "%b %Y"))
data$Month <- factor(data$Month, levels = unique(data$Month))

monthly_sale_type <- data %>% 
  group_by(Month, SaleType) %>% 
  summarise(Sales = sum(Price)) 

# Pivoring for better dataviz
data_wider = monthly_sale_type %>% 
  tidyr::pivot_wider(names_from = SaleType, values_from = Sales) 
## using more advanced functions from dplyr

all_monthly_metrics <- data %>% 
  group_by(Month, SaleType) %>% 
  summarise(across(Price,
                   list(min = min,
                        max = max,
                        mean = mean,
                        median = median)))
#.names = "{.fn}_price"


all_monthly_metrics_wider <- all_monthly_metrics %>% 
  tidyr::pivot_wider(names_from = SaleType, 
                     values_from = -c(1:2),
                     names_glue = '{SaleType}_{.value}') 


all_monthly_metrics_wider <- all_monthly_metrics_wider %>% 
  rename_with(~ stringr::str_remove(.x, "_Price"))

# preload styles ----

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
  fontColour = 'black',
  numFmt = "GENERAL",
  border = 'TopBottomLeftRight',
  borderColour = "darkgrey",
  borderStyle = "thin" ,
  bgFill = NULL,
  fgFill = '#FBFBFB',
  halign = 'left',
  valign = 'center')

wb <- openxlsx::createWorkbook()

sheet_i = 'Sales by Affiliate'

addWorksheet(wb, sheetName = sheet_i, tabColour = 'steelblue')

writeData(wb, sheet = sheet_i,
          x = data_wider, 
          headerStyle = my_headerStyle,
          borders = c('all'),
          borderColour = 'grey',
          borderStyle = 'thin',
          withFilter = TRUE)

addStyle(wb, sheet_i, style = my_bodyStyle,
         rows = 2:(1+nrow(data_wider)), cols = 1:ncol(data_wider),
         gridExpand = TRUE, stack = T)

date_style = createStyle(numFmt = "mmm yyyy",
                         wrapText = T,
                         halign = 'left',
                         textDecoration = 'bold')

addStyle(wb, sheet_i, style = date_style,
         rows = 2:(1+nrow(data_wider)), cols = 1,
         gridExpand = TRUE, stack = T)

# addStyle(wb, sheet_i,
#          style = createStyle(numFmt = 'CURRENCY'),
#          rows = 2:(1+nrow(data_wider)), cols = 2:4,
#          gridExpand = TRUE, stack = T)

setColWidths(wb, sheet_i, cols = 1:ncol(data_wider),  widths = 14)

openxlsx::openXL(wb)


sheet_i = 'All Monthly Metrics'

addWorksheet(wb, sheetName = sheet_i, tabColour = 'steelblue')

writeData(wb, sheet = sheet_i,
          x = all_monthly_metrics_wider, 
          headerStyle = my_headerStyle,
          borders = c('all'),
          borderColour = 'grey',
          borderStyle = 'thin',
          withFilter = TRUE)

addStyle(wb, sheet_i, 
         style = my_bodyStyle,
         rows = 2:(1+nrow(all_monthly_metrics_wider)), 
         cols = 1:ncol(all_monthly_metrics_wider),
         gridExpand = TRUE, stack = T)


addStyle(wb, sheet_i, 
         style = createStyle(numFmt = "#,##0.0"),
         rows = 2:(1+nrow(all_monthly_metrics_wider)), 
         cols = 2:ncol(all_monthly_metrics_wider),
         gridExpand = TRUE, stack = T)


openxlsx::openXL(wb)




# data %>%
#   group_by(Month, SaleType) %>%
#   summarise(
#     p10 = quantile(Price, 0.10, na.rm = TRUE),
#     p25 = quantile(Price, 0.25, na.rm = TRUE),
#     p50 = quantile(Price, 0.50, na.rm = TRUE),
#     p75 = quantile(Price, 0.75, na.rm = TRUE),
#     p90 = quantile(Price, 0.90, na.rm = TRUE)
#   )
