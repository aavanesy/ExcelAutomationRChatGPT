
library(dplyr)
library(lubridate)

start_date = as.Date('2020-01-01')
end_date = as.Date('2023-01-31')

nsales = 20

s1 <- sample(seq(start_date, end_date, by = 1), nsales, replace = T)
s2 <- sample(seq(start_date, end_date, by = 1), nsales, replace = T)
s3 <- sample(seq(start_date, end_date, by = 1), nsales, replace = T)
s4 <- sample(seq(start_date, end_date, by = 1), nsales, replace = T)
s5 <- sample(seq(start_date, end_date, by = 1), nsales, replace = T)

df <- data.frame(s1 = s1,
                 s2 = s2,
                 s3 = s3,
                 s4 = s4,
                 s5 = s5,
                 s6 = as.character(s1))

str(df)

# openxlsx code ----
library(openxlsx)
wb <- createWorkbook()
addWorksheet(wb, 'SheetDates')

writeData(wb, 'SheetDates', x = df)

openXL(wb)

addStyle(wb, 'SheetDates',
         style = createStyle(numFmt = "DATE"),
         rows = 2:21,
         cols = 6, gridExpand = TRUE) 

openXL(wb)


addStyle(wb, 'SheetDates',
         style = createStyle(numFmt = "yyyy/mm/dd"),
         rows = 2:21,
         cols = 1,
         gridExpand = TRUE) 

addStyle(wb, 'SheetDates',
         style = createStyle(numFmt = "yyyy/mmm/dd"),
         rows = 2:21,
         cols = 2,
         gridExpand = TRUE) 


addStyle(wb, 'SheetDates',
         style = createStyle(numFmt = "mm/dd/yyyy"),
         rows = 2:21,
         cols = 3,
         gridExpand = TRUE) 

addStyle(wb, 'SheetDates',
         style = createStyle(numFmt = "dddd"),
         rows = 2:21,
         cols = 4,
         gridExpand = TRUE) 


addStyle(wb, 'SheetDates',
         style = createStyle(numFmt = "dddd, mmmm d, yyyy"),
         rows = 2:21,
         cols = 5,
         gridExpand = TRUE) 


openXL(wb)

## Dates and Times ----
# from chatGPT ----
# Load necessary package
library(lubridate)

# Define the start and end date-time
start_datetime <- ymd_hms("2020-01-01 00:00:00")
end_datetime <- ymd_hms("2024-01-01 23:59:59")

# Generate random date-time values
generate_random_datetime <- function(n, start, end) {
  as_datetime(runif(n, as.numeric(start), as.numeric(end)), origin = "1970-01-01")
}

# Generate 10 random date-time values
random_datetimes <- generate_random_datetime(10, start_datetime, end_datetime)

# Print the random date-time values
print(random_datetimes)

df <- data.frame(s1 = random_datetimes,
                 s2 = random_datetimes,
                 s3 = random_datetimes,
                 s4 = random_datetimes)

str(df)

wb <- createWorkbook()
addWorksheet(wb, 'SheetTime')

writeData(wb, 'SheetTime', x = df)

openXL(wb)

addStyle(wb, 'SheetTime',
         style = createStyle(numFmt = 'mm/dd/yyyy hh:mm:ss'),
         rows = 2:11,
         cols = 2,
         gridExpand = TRUE) 

addStyle(wb, 'SheetTime',
         style = createStyle(numFmt = 'hh:mm:ss'),
         rows = 2:11,
         cols = 3,
         gridExpand = TRUE) 

addStyle(wb, 'SheetTime',
         style = createStyle(numFmt = 'dddd hh:mm:ss AM/PM'),
         rows = 2:11,
         cols = 4,
         gridExpand = TRUE) 


openXL(wb)
