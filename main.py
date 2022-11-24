import pandas
import pandas as pd
import os
import openpyxl
from openpyxl.styles import Font
from openpyxl import Workbook

os.remove('FinalTable.xlsx')
my_file = open("FinalTable.xlsx", "w+")
my_file.close()

pandas.read_json("gen.json").to_excel("output.xlsx")
path = "output.xlsx"

wb = Workbook()
sheet = wb.active

# Table header
c1 = sheet.cell(row = 1, column = 1)
c2 = sheet.cell(row = 1, column = 2)
c3 = sheet.cell(row = 1, column = 3)
c4 = sheet.cell(row = 1, column = 4)
c5 = sheet.cell(row = 1, column = 5)

c1.value = "IB"
c2.value = "class"
c3.value = "feature"
c4.value = "time"
c5.value = "value"

c1.font = Font(bold = True)
c2.font = Font(bold = True)
c3.font = Font(bold = True)
c4.font = Font(bold = True)
c5.font = Font(bold = True)

# Table header --end

df_len = len(pd.read_excel(path, "Sheet1"))

wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active

count_ib = 1
table_row = 2

for i in range(2, df_len + 2):

    cell_obj = sheet_obj.cell(row=i, column=3).value

    open_brackets = 0
    close_brackets = 0

    j = 1

    while j < len(cell_obj):

        if cell_obj[j] == '{':
            open_brackets += 1
            feature_name = ""
            j += 13

            while cell_obj[j] != "'":
                feature_name += cell_obj[j]
                j += 1

            while open_brackets != close_brackets:
                time = ""
                value = ""
                if cell_obj[j] == '{':
                    open_brackets += 1
                    j += 8
                    while cell_obj[j] != ",":
                        time += cell_obj[j]
                        j += 1

                    j += 12
                    while cell_obj[j] != "'":
                        value += cell_obj[j]
                        j += 1

                    c11 = sheet.cell(row=table_row, column=1)
                    c11.value = i - 1

                    c12 = sheet.cell(row=table_row, column=2)
                    class_name = sheet_obj.cell(row=i, column=2).value
                    c12.value = class_name

                    c13 = sheet.cell(row=table_row, column=3)
                    c13.value = feature_name

                    c14 = sheet.cell(row=table_row, column=4)
                    c14.value = time

                    c15 = sheet.cell(row=table_row, column=5)
                    c15.value = value

                    table_row += 1

                elif cell_obj[j] == '}':
                    close_brackets += 1
                j += 1

        j += 1

wb.save("FinalTable.xlsx")
