import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference

# File load karein
wb = openpyxl.load_workbook("Thiranex_Sales_Revenue_Analysis_Data.xlsx")
ws_data = wb["Sales Data"]
ws_dash = wb.create_sheet("Visual Dashboard")

# 1. Bar Chart: Revenue by Category
# (Maana ki data rows 200 tak hai)
chart1 = BarChart()
chart1.title = "Revenue by Category"
chart1.y_axis.title = 'Revenue'
chart1.x_axis.title = 'Category'

# G column me Revenue hai, B column me Category hai
data = Reference(ws_data, min_col=7, min_row=1, max_row=201)
cats = Reference(ws_data, min_col=2, min_row=2, max_row=201)

chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
ws_dash.add_chart(chart1, "A2")

# 2. Line Chart: Sales Trend
chart2 = LineChart()
chart2.title = "Sales Over Time"
chart2.style = 13 # Blue style
data_trend = Reference(ws_data, min_col=7, min_row=1, max_row=201)
chart2.add_data(data_trend, titles_from_data=True)
ws_dash.add_chart(chart2, "I2")

wb.save("Thiranex_Final_Dashboard.xlsx")
print("Dashboard Charts Generated Successfully!")