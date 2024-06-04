library(openxlsx)
library(dplyr)
library(lubridate)

# when reading a csv > mutate date
data <- read.csv('data/Sales_Shops.csv') %>% 
  head(100)

str(data)

data <- data %>% 
  mutate(Date = as.Date(Date))

str(data)

## without formatting, all-in-one ----

df1 = mtcars
df2 = iris
list.data <- list('df1' = df1, 'df2' = df2)
my_list <- list(df1, df2)

df3 = data.frame(x = 1, y = 2)

my_list <- c(my_list, list(df3))

openxlsx::write.xlsx(x = my_list, file = 'output/My_list.xlsx')

?split

data_list <- data %>% 
  split(f = as.factor(.$Country)) 
  
openxlsx::write.xlsx(x = data_list, file = 'output/DataPerCountry.xlsx')

data %>% 
  split(f = as.factor(.$Country)) %>% 
  openxlsx::write.xlsx(file = 'output/DataPerCountry.xlsx')

# basic formatting ----

# Define header style
my_headerStyle <- createStyle(
  fontSize = 12, 
  fontColour = "#FFFFFF", 
  halign = "center", 
  valign = "center", 
  fgFill = "#4F81BD", 
  border = "TopBottomLeftRight",
  borderColour = "#4F81BD",
  textDecoration = "bold"
)

data %>% 
  split(f = as.factor(.$Country)) %>% 
  openxlsx::write.xlsx(file = 'output/DataPerCountry2.xlsx',
                       headerStyle = my_headerStyle,
                       borders = c('all'),
                       borderColour = 'grey',
                       borderStyle = 'thin',
                       withFilter = TRUE)

# additional params ----

data %>% 
  split(f = as.factor(.$Country)) %>% 
  openxlsx::write.xlsx(file = 'output/DataPerCountry3.xlsx',
                       firstRow = T,
                       firstCol = T)


# saving as table

data %>% 
  split(f = as.factor(.$Country)) %>% 
  openxlsx::write.xlsx(file = 'output/DataPerCountryTable4.xlsx',
                       asTable = T,
                       tableStyle = 'TableStyleLight17',
                       overwrite = TRUE)



