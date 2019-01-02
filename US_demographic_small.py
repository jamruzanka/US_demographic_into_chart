from openpyxl import Workbook
from openpyxl.chart import Reference, Series, BarChart

class DataCharts(object):
    #To create an object, provide a txt's file name (existing one) and then the xlsx file (which will be created).
    def __init__(self, txt_file, excel_file_name):
        self.txt_file = txt_file
        self.excel_file_name = excel_file_name

    def txt_to_excel(self, split_symbol):
    #To convert a txt file into a xlsx file, provide a split symbol used in the txt file - could be ",", "|" etc.
        self.excel_file = Workbook()
        ws1 = self.excel_file.active

        file = open(self.txt_file, 'r')
        #Provided letters are for 12 columns now - if there is more columns, please add further letters as they appear in excel
        column_letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]
        row_number = 0
        column_number = 0
        for line in file:
            elements = line.split(split_symbol)
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
        ws1 = self.excel_file.active
        data = Reference(ws1, min_col = data_min_col, min_row = data_min_row, max_col = data_max_col, max_row = data_max_row)
        titles = Reference(ws1, min_col = ref_min_col, min_row = ref_min_row, max_row = ref_max_row)
        chart = BarChart()
        chart.title = chart_title
        chart.add_data(data=data, titles_from_data=True)
        chart.set_categories(titles)

        ws1.add_chart(chart, anchor_cell)

        self.excel_file.save(self.excel_file_name)

first_chart = DataCharts("US_data.txt", "US_data_class_script.xlsx")
first_chart.txt_to_excel("|")
#Anchor_cell is the cell where the chart will be created (it's top left corner).
#The arguments after the anchor_cell indicate the data that will be used to create a chart:
#data_min_col is the minimum column, data_min_row is the minimum row etc.).
#Without them, the script won't know which data should he visualize in the chart.
first_chart.create_chart("Birth and Internet", "H4", 2, 1, 3, 15, 1, 2, 15)
