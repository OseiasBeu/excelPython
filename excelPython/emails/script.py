import os
os.system('clear' if os.name == 'nt' else 'cls')

import pandas as pd
import numpy as np
from openpyxl import load_workbook

wb = load_workbook('emails.xlsx')

arquivo_excel  = wb.active
arquivo_excel.auto_filter.ref = "A:H"

arquivo_excel.auto_filter.add_filter_column(7,["jqualho@gmail.com"])
wb.save('emails.xlsx')