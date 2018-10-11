#! /usr/bin/env python3

import datetime
import xdrlib, sys, os, shutil
import xlrd


total_expcected = 0
total_moved = 0

fo = open("log.txt", "a") 

def open_xls(file_name):
	try:
		sheets = xlrd.open_workbook(file_name)		
		return sheets
	except Exception:
		log(0, "Oops! The file " + file_name + " is invalid.")
	
def assign_skus(sheets, sheet_name, path):	
	global total_expcected
	global total_moved
	src = path + "\\Todo\\"	
	level = ["Special1", "Special2", "Special3", "Complex1", "Complex2", "Complex3", "Medium1", "Medium2", "Medium3", "Simple"]
			
	src_folders = os.listdir(src)
	
	table = sheets.sheet_by_name(sheet_name)
	nrows = table.nrows
	for row in range(nrows):
		row_data = table.row_values(row)
		name = row_data[0]
		if name == "Name":
			continue
		if len(row_data) != 12:
			continue
		special_nu = int(row_data[2])
		complex_nu = int(row_data[3])
		medium_nu = int(row_data[4])
		simple_nu = int(row_data[5])
		
		total_expcected = 0
		total_moved = 0
		log(1, "Start assignment: " + name)
		dst_folder = path + "\\Assignment List\\"	 + name + "\\"
		
		for i in range(len(level)):
		    nu = int(row_data[i + 2])
		    cur_level = level[i]
		    mv_main(cur_level, nu, 0, src + cur_level, dst_folder + cur_level)
			
		log(1, "   ", "Total Expected:" , total_expcected , "  Total Assigned :", total_moved)
		print("")

def mv_main(level, nu, moved, src, dst_folder):
	global total_expcected
	global total_moved
	mv_list = []
	if nu != 0 :
		moved = mv_to_dst(mv_list, level, nu, moved, src, dst_folder)
		total_expcected += nu
		total_moved += moved
	log(1, "   ", level, "Expected:" , nu , " Assigned", ":", moved)
	print("        SKU:", mv_list)
		
def mv_to_dst(mv_list, level, nu, moved, src, dst_folder):
	
	if not os.path.exists(dst_folder):
		os.makedirs(dst_folder)
		
	log(0, "src", src)
	if not os.path.exists(src):
		log(0, "   ", level, "not found!")
		return moved
		
	sku_folders = os.listdir(src)
	count = 0
	for sku in sku_folders:	
		try:			
			if is_sku(src, sku):
				shutil.move(src + "\\" + sku, dst_folder)	
				# log(0, src + "\\" + sku, dst_folder)
				moved += 1	
				log(0, "moved increase for", sku, moved)
				mv_list.append(sku)
			else:
				log(0, "check another folder")
				moved = mv_to_dst(mv_list, level, nu, moved, src + "\\" + sku, dst_folder)
				if moved == nu:
					log(0, "breakout")
					return moved
				log(0, "check another folder end", count, moved)
							
			if moved == nu:
				log(0, "breakout")
				return moved			
		except Exception as e:	
			log (1, "Move failed for" , sku , e)	
	
	return moved
	
def log(debug=0, *args):
	if debug == 0:
		pass
		# print(*args)
	else:
		print(*args)
		for arg in args:
			fo.write(str(arg))
		fo.write("\r")
	
def is_sku(src, sku):	
	result = False
	sku_folders = os.listdir(src + "\\" + sku)
	if len(sku_folders) == 0:
		log(0, src, sku, "is empty")
		return False
	for folder in sku_folders:
		log(0, folder, "check")
		if os.path.isdir(src + "\\" + sku + "\\" + folder):
			log(0, src, sku, "is not sku")
			return False
		if folder.find("jpg") > -1:
			result = True
			
	log(0, src, sku, "is sku")
	return result
	
def main():	
	# team = sys.argv[1]
	file_name = "Team task assignment"	
	promp = False
	if len(sys.argv) == 1:
		promp = True
		sheet_name = input("Input excel task:")
		
	if len(sys.argv) == 2:
		sheet_name = sys.argv[1]	
		
	now = datetime.datetime.now()
	now = now.strftime('%Y-%m-%d %H:%M:%S')
	log(1, "\n", now, " Start assign ", sheet_name, "\n")
	path = os.getcwd()			
	sheets = open_xls(path + "\\" + file_name + ".xlsx")		
	
	if sheets:	
		assign_skus(sheets, sheet_name, path)
	
	if promp:
		input("Press the enter key to exit.")
	
if __name__ == "__main__":
	main()	
