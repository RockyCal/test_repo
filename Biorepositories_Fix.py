# Clean up spreadsheet to fit C4P format
from openpyxl import load_workbook, cell
from openpyxl.cell import coordinate_from_string, column_index_from_string
import requests
from datetime import datetime, date, time

wb = load_workbook('biorepositories-csv.xlsx')
ws = wb.get_active_sheet()

START_ROW = 2
END_ROW = ws.get_highest_row()

class Entry(self, name):
	self.name = name

for row in ws.range('%s%s:%s%s'%(TITLE_COL, START_ROW, LAST_COL, END_ROW)):
	