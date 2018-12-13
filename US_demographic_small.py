from openpyxl import Workbook
from openpyxl.chart import Reference, Series, BarChart

my_excel = Workbook()
ws1 = my_excel.active
ws2 = my_excel.create_sheet("Sheet_2")

my_file = open('US_data.txt', 'r')
column_letters = ["A", "B", "C", "D", "E", "F", "G"]
row_number = 0
column_number = 0
for line in my_file:
    elements = line[:-1].split("|")
    if len(elements) > 1:
        for element in elements:
            try:
                floatElement = float(element)
                ws1[column_letters[column_number]+str(row_number + 1)] = floatElement
            except:
                ws1[column_letters[column_number]+str(row_number + 1)] = element
            column_number += 1
        row_number += 1
        column_number = 0

data = Reference(ws1, min_col = 2, min_row = 1,
                      max_col = 3, max_row = 15)
titles = Reference(ws1, min_col = 1, min_row = 2, max_row = 15)
US_chart = BarChart()
US_chart.title = "Birth rate and Internet users in US - percentage"
US_chart.add_data(data=data, titles_from_data=True)
US_chart.set_categories(titles)

ws1.add_chart(US_chart, "H3")

my_excel.save("US_demographic_small.xlsx")
