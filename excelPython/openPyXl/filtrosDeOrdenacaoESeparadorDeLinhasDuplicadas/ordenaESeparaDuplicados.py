import openpyxl

wb = openpyxl.load_workbook('test.xlsx')
ws1 = wb.active


names = []
for row in ws1.columns[0]:
    names.append(row.value)

names = sorted(list(set(names)))

start = 1
for name in names:
    ws1.cell(row = start, column=2).value = name
    start += 1

wb.save('test.xlsx')