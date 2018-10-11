#! /usr/bin/env python3

import xdrlib, sys, os, shutil
import xlrd
from xlutils.copy import copy
import xlwt

table = None

def open_xls(file_name):	
	try:
		data = xlrd.open_workbook(file_name)		
		return data
	except Exception:
		print("Oops! The file " + file_name + " is invalid.")
	
def read_sheet(data, index):
	rtable = data.sheet_by_name(index)	
	wb = copy(data)
	wrtable = wb.get_sheet(index)
	return (rtable, wrtable, wb)
	
def fill_data(src, table):	
	src_folders = os.listdir(src)
	for name in src_folders:	
		if is_sku(src, name):
			fill_to_excel(src, name, table)
		else:
			if os.path.isdir(src + "\\" + name):
				fill_data(src + "\\" + name, table)
		

def fill_to_excel(src, name, table):	
	#global table
	#col = 24
	#ctype = 1
	#table.put_cell(row, col, ctype, value)
	rtable = table[0]
	wtable = table[1]
	nrows = rtable.nrows
	for row in range(nrows):
		val = rtable.row_values(row)	
		sku_nu = rtable.row(row)[3].value
		if name == sku_nu.strip():
			#print(sku_nu, row)
			fill_to_row(row, src, name, wtable)
			break
	
def fill_to_row(row, src, name, wtable):
	ctype = 1
	xf_index = 0
	
	index = src.find("打等级")
	if index > -1:
		#print(src, index, src[index+4:])
		src = src[index+4:]
		src = src.replace("\\可变色", "");
		level = src.split('\\')
		if len(level) < 2:
			print("Error input for:", name)
			return
		category = level[0]
		complexity = get_complexity(level[1:])
		similar = get_similar(level[1:], category)

		print(name,row, complexity, similar)	
				
		wtable.write(row, 0, category)
		wtable.write(row, 24, name)
		wtable.write(row, 25, complexity)
		wtable.write(row, 26, similar)
	

def get_complexity(level):
	for l in level:
		if l.find("相似") == -1:
			return l
	return ""
	
def get_similar(level, category):
	for l in level:
		if l.find("相似") > -1:
			return l
	return category
	
def is_sku(src, sku):	
	result = False
	sku_folders = os.listdir(src + "\\" + sku)
	if len(sku_folders) == 0:
		return False
	for folder in sku_folders:
		if os.path.isdir(src + "\\" + sku + "\\" + folder):
			return False
		if folder.find("jpg") > -1:
			result = True
			
	return result

def save_excel(path, name, table):
	wb = table[2]
	name = name.replace(".xlsx", ".xls");
	name = name.replace(".xls", "_Done.xls");
	if name.find("xls") == -1:
		name = name + "_Done.xls"
	wb.save(path + "\\" + name)
	
def main():	
	
	file = "ROBatch0006_Test.xlsx"	
	if len(sys.argv) > 1:
		file = sys.argv[1]
		
	path = os.getcwd()		
	raw_data = open_xls(path + "\\" + file)		
		
	table = read_sheet(raw_data, "SKU List")	
	
	fill_data(path+"\\打等级", table)
	
	save_excel(path, file, table)
	
	
if __name__ == "__main__":
	main()	