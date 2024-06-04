
library(dplyr)
library(lubridate)

start_date = as.Date('2020-01-01')
end_date = as.Date('2023-01-31')

setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
getwd()

shops = c('Amazon', 'Wallmart', 'Costco')
country = c('US', 'Canada', 'UK', 'France')
items = c('Keyboard', 'Mouse', 'UsbCable', 'Charger')
sale_type = c('Direct', 'Promotion', 'Affiliate')
price_range = 20:200 / 10

nsales = 100

table_res = tibble(Date = sample(seq(start_date, end_date, by = 1), nsales, replace = T),
                   Shop = sample(shops, nsales, replace = T),
                   Item = sample(items, nsales, replace = T),
                   Country = sample(country, nsales, replace = T),
                   SaleType = sample(sale_type, nsales, replace = T),
                   Price = sample(price_range, nsales, replace = T),
                   PriceDiscount = Price * 0.9,
                   ClickRate = sample(1:10/100, nsales, replace = T),
                   URL = 'https://copilot.microsoft.com/')
str(table_res)

class(table_res$Price) <- "currency"
class(table_res$PriceDiscount) <- "accounting"
class(table_res$ClickRate) <- "percentage"
class(table_res$URL) <- "hyperlink"

wb <- createWorkbook()
addWorksheet(wb, sheetName = 'TableOne')

?writeDataTable

writeDataTable(wb, 'TableOne', table_res, 
               tableStyle = "TableStyleMedium9")

openXL(wb)


openxlsx::saveWorkbook(wb, 'data/datatable.xlsx', overwrite = T)



