import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

#----------------------------------------------------
def excel_merge (writer, xls1, xls2, sheet, index=None, sort=None):
	df_xls1 = pd.read_excel(xls1, sheet).fillna("#-")
	df_xls2 = pd.read_excel(xls2, sheet).fillna("#-")

	if index != None:
		df_xls1 = df_xls1.set_index(index)
		df_xls2 = df_xls2.set_index(index)
	
	df = pd.concat([df_xls1, df_xls2], axis=1).fillna("X")
	df = df.replace(to_replace="#-", value=np.nan)

	if sort != None:
		df = df.sort_values(by=sort, ascending=True)

	df.to_excel(writer, sheet_name=sheet, startrow = 1)
	headers(writer.sheets[sheet], sheet)
	formatting(writer.sheets[sheet], sheet)
	#writer.save()
#----------------------------------------------------
def headers (ws, sheet):
	ws["B1"] = "Nexus A"
	font = Font(name = "Calibri", size = 12, bold = True)
	alignment = Alignment(horizontal="center", vertical="center")
	border = Border()
	ws["B1"].font = font
	ws["B1"].alignment = alignment
	if sheet == "VLANs":	
		ws["C1"] = "Nexus B"
		ws["C1"].font = font
		ws["C1"].alignment = alignment
	if sheet == "SVIs":
		ws["F1"] = "Nexus B"
		ws["F1"].font = font
		ws["F1"].alignment = alignment
		ws.merge_cells(start_row = 1, start_column = 2, end_row = 1, end_column = 6)
		ws.merge_cells(start_row = 1, start_column = 7, end_row = 1, end_column = 10)
	if sheet == "Ints":
		ws["H1"] = "Nexus B"
		ws["H1"].font = font
		ws["H1"].alignment = alignment
		ws.merge_cells(start_row = 1, start_column = 2, end_row = 1, end_column = 8)
		ws.merge_cells(start_row = 1, start_column = 9, end_row = 1, end_column = 14)
	if sheet == "Po":
		ws["G1"] = "Nexus B"
		ws["G1"].font = font
		ws["G1"].alignment = alignment
		ws.merge_cells(start_row = 1, start_column = 2, end_row = 1, end_column = 7)
		ws.merge_cells(start_row = 1, start_column = 8, end_row = 1, end_column = 13)
#----------------------------------------------------
def formatting (ws, sheet):
	i = 0
	cmax = [6] * ws.max_column
	max_column = cmax
	thin = Side(border_style = "thin", color = "000000")
	#for r in ws["A1:"+chr(ws.max_column+64)+str(ws.max_row)]:
	for r in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
		for c in r:
			c.border = Border(bottom = thin, top = thin, right = thin, left = thin)
			if c.value == "X":
				c.fill = PatternFill(start_color = "FB2C57", end_color = "FB2C57", fill_type = "solid")
#Ajuste del ancho de columnas
			try:
				lenght = len(c.value)
			except TypeError as uni:
				lenght = 0
			cmax[i] = max(lenght, cmax[i])
			i += 1
		i = 0
		for j in range(ws.max_column):
			max_column[j] = max(max_column[j], cmax[j]) 
	for j in range(len(max_column)):
		ws.column_dimensions[get_column_letter(j+1)].width = max_column[j]
#----------------------------------------------------
def write_excel (path1, path2):
	xls1 = pd.ExcelFile(path1)
	xls2 = pd.ExcelFile(path2)
	wb = openpyxl.Workbook()
	with pd.ExcelWriter("conf_merge.xlsx", engine="openpyxl") as writer:
		excel_merge (writer, xls1, xls2, "VLANs", "VLAN", "VLAN")
		excel_merge (writer, xls1, xls2, "SVIs", "SVI", "SVI")
		excel_merge (writer, xls1, xls2, "Ints", "Interface", "Interface")
		excel_merge (writer, xls1, xls2, "Po", "Interface", "Interface")
		excel_merge (writer, xls1, xls2, "Statics")
		excel_merge (writer, xls1, xls2, "IP_ACLs")
#----------------------------------------------------

if __name__ == "__main__":
	import tkinter as tk
	from tkinter import filedialog
	root = tk.Tk()
	root.withdraw()
	path1 = filedialog.askopenfilename()
	path2 = filedialog.askopenfilename()
	write_excel (path1, path2)




