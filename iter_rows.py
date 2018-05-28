import os
from openpyxl import *
import pprint



wb = load_workbook(filename="balances.xlsx", read_only=True,data_only=True)
ws = wb["Sheet"]

for row in ws.iter_rows(min_row=7):
	if row[0].value:
		print row[0].value
	else:
		break
