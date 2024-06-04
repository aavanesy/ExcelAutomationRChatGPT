
library(openxlsx2)
library(mschart)


# code scripts ----
wb <- wb_workbook()
wb$add_worksheet("add_mschart")

?ms_linechart

chart_01 <- ms_linechart(
  data = us_indus_prod,
  x = "date", y = "value",
  group = "type"
)

chart_01 <- chart_ax_x(
  x = chart_01, 
  num_fmt = "mmm yyyy",
  limit_min = min(us_indus_prod$date), 
  limit_max = as.Date("1992-01-01")
)
chart_01 = chart_data_line_width(x = chart_01, 1)


wb <- wb %>%
  wb_add_mschart(
    dims = "D2:H20",
    graph = chart_01)

openxlsx2::wb_open(wb)


openxlsx2::wb_save(wb, 'output/mschart.xlsx')