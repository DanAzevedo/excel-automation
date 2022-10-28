from openpyxl import Workbook
from openpyxl.chart import (
    LineChart,  # Permite criar gráficos
    Reference,  # Vincula e linkar de onde o linechart vai consumir os seus dados
)
from datetime import date

wb = Workbook()
ws = wb.active

rows = [
    ['Data', 'Batch 1', 'Batch 2', 'Batch 3'],
    [date(2015, 9, 1), 40, 30, 25],
    [date(2015, 9, 2), 40, 25, 30],
    [date(2015, 9, 3), 50, 30, 45],
    [date(2015, 9, 4), 30, 25, 40],
    [date(2015, 9, 5), 25, 35, 30],
    [date(2015, 9, 6), 20, 40, 35],
]

# Outra forma de adicionar dados dentro da sheet sem ter que pegar célula por célula
for row in rows:
    ws.append(row)

c1 = LineChart()
c1.title = "Line Chart"
c1.y_axis.title = "Size"
c1.x_axis.tile = "Número de teste"

data = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=7)
c1.add_data(data, titles_from_data=True)

c1.series[0].marker.symbol = 'triangle'
c1.series[0].marker.graphicalProperties.solidFill = 'FF0000'

ws.add_chart(c1, "A10")

wb.save('line.xlsx')
