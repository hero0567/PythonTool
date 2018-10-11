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
		
	table = data.sheet_by_name(index)	
	nrows = table.nrows
	list = []
	for row in range(nrows):
		val = table.row_values(row)		
		list.extend(val)	
	return list
	
def create_models(path, list, date):	
	src = path+"\\Model deliverables"
	path =path+"\\Deliver file" #change created folder name
	src_folders = os.listdir(src)
	total=0
	incorrect=0
	for name in src_folders:
		for sku in list:
			sku = str(int(sku))
			if name.find(sku) > -1:
				dst = path + "\\" + date + "\\" + sku
				print("Find sku:" + dst)
				total+=1
				shutil.copytree(src + "\\" + name, dst)		
								
				dst_folders = os.listdir(dst)								
				for folder_name in dst_folders:	
					if folder_name.find(".zip") > -1:
						print("Folder Name:" + folder_name)
						incorrect+=1
					if len(dst_folders) == 1:						
						if len(folder_name) > 8:
							os.rename(dst + "\\" + folder_name, dst + "\\" + folder_name[0:8])							
						break
					shutil.rmtree(dst + "\\" + folder_name)
					dst_folders = os.listdir(dst)				
	print("Copy Total:" + str(total))
	print("Incorrect Total:" + str(incorrect))

def remove_cache_files(src):	
	src_folders = os.listdir(src)
	for name in src_folders:
		if os.path.isdir(src + "\\" + name):
			remove_cache_files(src + "\\" + name)
		else:
			if name == "Thumbs.db":
				os.remove(src + "\\" + name)
				
				
def find_miss_file(sku_list, des):
	totol = 0
	missing = False
	src_folders = os.listdir(des)
	for sku in sku_list:
			sku = str(int(sku))			
			if exist_zip(des + "\\" + sku) == 0:
				missing = True
				totol += 1
				print("Mssing zip:   ", sku)
			if exist_zip(des + "\\" + sku) == 1:
				missing = True
				totol += 1
				print("Mssing folder:", sku)
				
	if not missing:
		print("No missing sku found!")	
	else:
		print("Totol missing:", totol)

def exist_zip(des):
    # 0 missing zip, 1 missing folder, 2 OK
	existed = 0
	if not os.path.exists(des):
		return 1
	
	folders = os.listdir(des)
	for folder in folders:
		if os.path.isdir(des + "\\" + folder):
			existed = exist_zip(des + "\\" + folder)
		else:
			if folder.find(".zip") > -1:
				existed = 2
				break
	return existed

	
def main():	
	date = sys.argv[1]
		
	path = os.getcwd()
		
	raw_data = open_xls(path + "\\QACustomer.xls")		
		
	sku_list = read_sheet(raw_data, date)	
	
	remove_cache_files(path+"\\Model deliverables")
	create_models(path, sku_list, date)
	find_miss_file(sku_list, path+"\\Deliver file" + "\\" +date)
	
if __name__ == "__main__":
	main()	