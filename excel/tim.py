from openpyxl import Workbook, load_workbook


wb = Workbook()
ws = wb.active
ws.title = "Data"



ws.append(['Tim', 'Is', 'Great'])

wb.save('tim.xlsx')