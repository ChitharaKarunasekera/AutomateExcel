from openpyxl import Workbook, load_workbook

wb = Workbook()  # initialize new workbook
ws = wb.active  # Access sheet. The default sheet
ws.title = "Data"

# inserting many data at once
ws.append(['Chithara', 'Is', 'Great', '!'])

wb.save('Chithara.xlsx')  # saving file
