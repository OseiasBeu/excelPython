import os
os.system('clear' if os.name == 'nt' else 'cls')

from openpyxl import load_workbook
caminho = 'PONTO.xlsx'
arquivo_excel = load_workbook(caminho)

# Copiando dados de uma planilha para a outra:
original = arquivo_excel.get_sheet_by_name('Gastos.xlsx')
copia = arquivo_excel.copy_worksheet(original)

arquivo_excel.save('planilhaCopia.xlsx')

