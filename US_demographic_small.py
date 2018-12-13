from openpyxl import Workbook
from openpyxl.chart import Reference, Series, BarChart

class DataCharts(object):
    def __init__(self, txt_file, excel_file):
        self.txt_file = txt_file
        self.excel_file_name = excel_file

    def txt_to_excel(self):
        self.excel_file = Workbook()
        ws1 = self.excel_file.active

        file = open(self.txt_file, 'r')
        column_letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
        row_number = 0
        column_number = 0
        for line in file:
            elements = line.split("|")
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
        self.excel_file.save(self.excel_file_name)

    def create_chart(self, chart_title, anchor_cell, data_min_col, data_min_row, data_max_col, data_max_row, ref_min_col, ref_min_row, ref_max_row):
        #excel_file_open = self.excel_file.load_workbook()
        ws1 = self.excel_file.active
        data = Reference(ws1, min_col = data_min_col, min_row = data_min_row, max_col = data_max_col, max_row = data_max_row)
        titles = Reference(ws1, min_col = ref_min_col, min_row = ref_min_row, max_row = ref_max_row)
        chart = BarChart()
        chart.title = chart_title
        chart.add_data(data=data, titles_from_data=True)
        chart.set_categories(titles)

        ws1.add_chart(chart, anchor_cell)

        self.excel_file.save(self.excel_file_name)

first_chart = DataCharts("US_data.txt", "US_data_class_sctipt.xlsx")
first_chart.txt_to_excel()
first_chart.create_chart("Birth and Internet", "H4", 2, 1, 3, 15, 1, 2, 15)
