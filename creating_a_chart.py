import openpyxl
from openpyxl.chart import PieChart,Reference,Series,PieChart3D

wb = openpyxl.Workbook()
ws = wb.active

data = [
    ["IceCream","Sold"],
    ["Vanila",1500],
    ["Choclate",1100],
    ["ButterScotch",2500],
    ["Horlix",1200],
    ["Magnum",500],
    ["Pesta",500]
    ]


for i in data:
    ws.append(i)

chart = PieChart()
labels = Reference(ws, min_col = 1, min_row = 2, max_col = 5)
data = Reference(ws, min_col = 2, min_row = 1, max_row = 5)
chart.add_data(data,titles_from_data = True)
chart.set_categories(labels)
chart.title = "IceCream"



ws.add_chart(chart, 'c1')
wb.save("C:/Users/Mehedi Hassan Galib/Desktop/Python/chart.xlsx")
