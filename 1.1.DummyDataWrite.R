
library(dplyr)
library(lubridate)

start_date = as.Date('2020-01-01')
end_date = as.Date('2023-01-31')

getwd()

setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

getwd()

shops = c('Amazon', 'Wallmart', 'Costco')

country = c('US', 'Canada', 'UK', 'France')

items = c('Keyboard', 'Mouse', 'UsbCable', 'Charger')

sale_type = c('Direct', 'Promotion', 'Affiliate')

price_range = 20:200 / 10

nsales = 10000

table_res = tibble(Date = sample(seq(start_date, end_date, by = 1), nsales, replace = T),
                   Shop = sample(shops, nsales, replace = T),
                   Item = sample(items, nsales, replace = T),
                   Country = sample(country, nsales, replace = T),
                   SaleType = sample(sale_type, nsales, replace = T),
                   Price = sample(price_range, nsales, replace = T))


table_res = table_res %>% 
  arrange(Date)

str(table_res)

# saving in different formats

?write.csv
write.csv(table_res, file = paste0('data/Sales_Shops.csv'), 
          row.names = F)

?write.table
write.table(table_res, file = paste0('data/Sales_Shops.txt'), 
          row.names = F)

?openxlsx::write.xlsx
openxlsx::write.xlsx(x = table_res,
                     file = 'data/Sales_Shops.xlsx',
                     overwrite = TRUE)

# Example of writing numbers as character
table_res2 = table_res %>% 
  mutate(Price = as.character(Price))

str(table_res2)

openxlsx::write.xlsx(x = table_res2,
                     file = 'data/Sales_Shops_char_numb.xlsx',
                     overwrite = TRUE)
