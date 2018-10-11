#! /usr/bin/env python3

import xdrlib, sys, os, shutil
import xlrd

def open_xls(file_name):
	try:
		data = xlrd.open_workbook(file_name)		
		return data
	except Exception:
		print("Oops! The file " + file_name + " is invalid.")
	
def read_sheet(data, index):			
	# table = data.sheet_by_name(index)	
	table = data.sheet_by_index(0)	
	nrows = table.nrows
	list = []
	for row in range(nrows):
		val = table.row_values(row)		
		list.extend(val)			
	
	return list
					
def copy_skus(list, team):	
	src = "D:\\New order\\ROBatch0003_img\\imgs"
	dst = "D:\\New order\\ROBatch0003_img\\Assginment" + team + "\\Todo"		
		
	src_folders = os.listdir(src)
	
	j = 0
	for src_folder in src_folders:							
		length = len(list)	
		i = 0		
		while(length > 0):
			if src_folder.find(list[i]) > -1:					
				dst_folder = dst + "\\" + list[i+1] + "\\" + list[i]				
				try:
					shutil.copytree(src + "\\" + src_folder, dst_folder)					
					j += 1
					print(j, list[i], "copied.")					
					break
				except Exception:					
					continue	
			else:
				length -= 2
				i += 2
				
def main():	
	#team = sys.argv[1]
	team = "Mini"
		
	path = os.getcwd()	
		
	raw_data = open_xls(path + "\\" + team + ".xlsx")		
		
	sku_list = read_sheet(raw_data, team)		
	
	copy_skus(sku_list, team)
	
if __name__ == "__main__":
	main()	