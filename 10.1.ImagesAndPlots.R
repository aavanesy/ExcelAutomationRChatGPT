
library(openxlsx)
library(dplyr)
library(ggplot2)

wb <- createWorkbook()
addWorksheet(wb, sheet = 'Report')

setwd(dirname(rstudioapi::getActiveDocumentContext()$path))


# Insert Image ----

?openxlsx::insertImage

img = 'data/GoogleLogo.JPG'

# insertImage(wb, "Report",
#             img,
#             startRow = 1, 
#             startCol = 1, 
#             width = 6,
#             height = 5)

insertImage(wb, "Report",
            img,
            startRow = 1, 
            startCol = 1, 
            width = 6,
            height = 6/2.807)


insertImage(wb, "Report", img,
            startRow = 1, 
            startCol = 10, 
            width = 10,
            height = 10/2.807,
            units = "cm")

insertImage(wb, "Report", 
            img,
            startRow = 20, 
            startCol = 6, 
            width = 6,
            height = 6/2.807,
            dpi = 1200,
            units = "cm")

openXL(wb)

## Insert plot ----

?openxlsx::insertPlot

wb <- createWorkbook()
addWorksheet(wb, sheet = 'Plot', gridLines = F)

data <- read.csv('data/Sales_Shops.csv')

data <- data %>% 
  mutate(Date = as.Date(Date))

ugly_plot <- data %>%
  group_by(Country, Item) %>%
  summarise(AveragePrice = mean(Price, na.rm = TRUE)) %>% 
  ggplot(aes(x = Country, y = AveragePrice, fill = Item)) + 
  geom_col(position=ggplot2::position_dodge())

summary_data <- data %>%
  group_by(Country, Item) %>%
  summarize(Total_Sales = sum(Price))

# Create the bar chart
nice_plot <- ggplot(summary_data, aes(x = Item, y = Total_Sales, fill = Country)) +
  geom_bar(stat = "identity", position = "dodge") +
  labs(title = "Total Sales by Country and Item",
       x = "Item",
       y = "Total Sales (in dollars)") +
  theme_minimal() + 
  theme(
    panel.border = element_blank(),  # Remove any existing panel border
    plot.background = element_rect(colour = "grey50", fill = NA, size = 0.5)  # Add border to the exterior of the plot
  )

print(ugly_plot)

insertPlot(wb,
           1,
           # width = 5,
           # height = 3.5, 
           fileType = "png", 
           units = "in")

print(nice_plot)

insertPlot(wb,
           1,
           startCol = 10,
           startRow = 20,
           # width = 5,
           # height = 3.5, 
           fileType = "png", 
           units = "in")

openXL(wb)



