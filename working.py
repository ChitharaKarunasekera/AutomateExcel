from openpyxl import Workbook, load_workbook

# Section 1
wb = Workbook()  # initialize new workbook
ws = wb.active  # Access sheet. The default sheet
ws.title = "Data"

# inserting many data at once
ws.append(['Chithara', 'Is', 'Great', '!'])
ws.append(['Chithara', 'Is', 'Great', '!'])
ws.append(['Chithara', 'Is', 'Great', '!'])
ws.append(['Chithara', 'Is', 'Great', '!'])
ws.append(['Chithara', 'Is', 'Great', '!'])
ws.append(['End!'])

wb.save('Chithara.xlsx')  # saving file

# Section 2
wb = load_workbook('Chithara.xlsx')  # load workbook
ws = ws.active

# access row 1 through 10
for row in range(1, 11):
    for col in range(0, 4):
        # chr takes an int and gives the character representation of it
        char = chr(65 + col)
