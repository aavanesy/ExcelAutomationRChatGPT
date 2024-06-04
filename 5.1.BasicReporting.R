library(openxlsx)
library(dplyr)
library(lubridate)

setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

# preload styles ----

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
  # numFmt = "GENERAL",
  border = 'TopBottomLeftRight',
  borderColour = "darkgrey",
  borderStyle = "thin" ,
  bgFill = NULL,
  fgFill = '#FBFBFB',
  halign = 'left',
  valign = 'center')

# data analytics ----

# when reading a csv > mutate date
data <- read.csv('data/Sales_Shops.csv')

str(data)

data <- data %>% 
  mutate(Date = as.Date(Date))

?as.Date
as.Date('01/15/2020', "%m/%d/%Y")
# as.Date('01/15/2020', "%m/%d/%y")
as.Date('15/01/2020', "%d/%m/%Y")
?lubridate::mdy

mdy('01/15/2020')
dmy('15/01/2020')

str(data)

# when reading from excel 
data <- read.xlsx(xlsxFile = 'data/Sales_Shops.xlsx', sheet = 1)

str(data)

# as.Date(43831, origin = "1899-12-30")

data$Date[1]

?openxlsx::convertToDate
openxlsx::convertToDate(data$Date[1])

data <- data %>% 
  mutate(Date = convertToDate(Date))

## Some analytics with ChatGPT

total_sales_per_shop <- data %>%
  group_by(Shop) %>%
  summarise(TotalSales = sum(Price, na.rm = TRUE))

print(total_sales_per_shop)


data %>%
  group_by(Shop, Item) %>%
  summarise(TotalSales = sum(Price, na.rm = TRUE))

average_price_per_item <- data %>%
  group_by(Item) %>%
  summarise(AveragePrice = mean(Price, na.rm = TRUE))

print(average_price_per_item)

highest_sale_per_country <- data %>%
  group_by(Country) %>%
  summarise(HighestSale = max(Price, na.rm = TRUE))

print(highest_sale_per_country)

summary_report <- data %>%
  group_by(Shop, Country, Item, SaleType) %>%
  summarise(
    TotalSales = sum(Price, na.rm = TRUE),
    AveragePrice = mean(Price, na.rm = TRUE),
    Count = n(),
    HighestSale = max(Price, na.rm = TRUE)
  ) %>%
  ungroup()

print(summary_report)




## grouping and summarize ----

annual_sales_shop <- data %>% 
  mutate(Year = year(Date)) %>% 
  group_by(Year, Shop) %>% 
  summarise(Sales = sum(Price))

class(annual_sales_shop$Sales) = 'accounting'

wb <- openxlsx::createWorkbook()

sheet_i = 'Summary'

addWorksheet(wb, sheetName = sheet_i, tabColour = 'steelblue')

writeData(wb, 
          sheet = sheet_i,
          x = annual_sales_shop, 
          headerStyle = my_headerStyle,
          borders = c('all'),
          borderColour = 'grey',
          borderStyle = 'thin',
          withFilter = TRUE)

openxlsx::openXL(wb)

style_accounting <- createStyle(numFmt = 'ACCOUNTING')

addStyle(wb, sheet_i,
         style = my_bodyStyle,
         rows = 2:(1+nrow(annual_sales_shop)),
         cols = 1:ncol(annual_sales_shop),
         gridExpand = TRUE)

addStyle(wb, sheet_i,
         style = style_accounting,
         rows = 2:(1+nrow(annual_sales_shop)),
         cols = 3,
         gridExpand = TRUE, stack  = T)

openxlsx::openXL(wb)


sheet_x = 'ChatGPTReport'
addWorksheet(wb, sheetName = sheet_x, tabColour = 'orange')

writeData(wb, 
          sheet = sheet_x,
          x = summary_report, 
          headerStyle = my_headerStyle,
          borders = c('all'),
          borderColour = 'grey',
          borderStyle = 'thin',
          withFilter = TRUE)

addStyle(wb, sheet_x,
         style = style_accounting,
         rows = 2:(1+nrow(summary_report)),
         cols = 5:8,
         gridExpand = TRUE, stack  = T)


sheet_j = 'Raw Data'

addWorksheet(wb, sheetName = sheet_j, tabColour = 'grey')

writeData(wb, sheet = sheet_j,
          x = data, 
          headerStyle = my_headerStyle,
          borders = c('all'),
          borderColour = 'grey',
          borderStyle = 'thin',
          withFilter = TRUE)



saveWorkbook(wb, 'output/Report.xlsx', overwrite = T)


## monthly report chatgpt

df <- data %>%
  mutate(Month = format(Date, "%b %Y"))

# Convert Month column to factor and specify the levels to ensure correct ordering
df$Month <- factor(df$Month, levels = unique(df$Month))

# Group by Month, SaleType, and Country and calculate total sales
monthly_sales_report <- df %>%
  group_by(Month, SaleType, Country) %>%
  summarise(TotalSales = sum(Price))

# Print the monthly sales report
print(monthly_sales_report)

monthly_sales_report %>%
  ggplot(aes(x = Month, y = TotalSales, color = Country, group = Country)) +
  geom_point() +
  facet_grid(~SaleType)
