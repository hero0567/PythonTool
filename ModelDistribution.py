#! /usr/bin/env python3

import xdrlib, sys, os, shutil
import xlrd

fo = open("log.txt", "w")

def open_xls(file_name):
	try:
		data = xlrd.open_workbook(file_name)		
		return data
	except Exception:
		print("Oops! The file " + file_name + " is invalid.")
	
def read_sheet(data, name):	
		
	table = data.sheet_by_name(name)		
	return table
	
def find_sku_zip(path, table):	
	src = path+"\\Model Deliverables"
	total=0
	not_found=0
	copy_failed=0
	nrows = table.nrows
	
	for row in range(nrows):
		found = False
		row_data = table.row_values(row)
		sku = row_data[0]
		type = row_data[1]
		sku = str(int(row_data[0]))
		zip_path = find_one_sku_zip(src, sku)
		if zip_path != None:
			copied = copy_zip_to_dest(zip_path, sku, type)
			if copied:				
				total+=1
				found = True
				log(1, "Copy zip:", sku)
			else:
				copy_failed+=1
				log(1, "Copy failed：", sku)
		else:
			log(1, "Not found：", sku)
			not_found+=1								
	log(1, "Zip copy Total:" , str(total))
	log(1, "Zip copy failed  Total:" , str(copy_failed))
	log(1, "Zip not found  Total:" , str(not_found))
	
def find_one_sku_zip(path, sku):
	src_folders = os.listdir(path)
	for name in src_folders:
		new_path = path + "\\" + name
		if os.path.isdir(new_path):
			found = find_one_sku_zip(new_path, sku)
			if found:
				return found				
		else:
			if name.find(sku) > -1 and name.find(".zip") > -1:
				#find zip file, but need get the newest one
				list = path.split("\\")
				new_list = list[0:len(list)-1]
				new_path = "\\".join(new_list)
				return new_path
	return None	
	
	
def find_sku_jpg(src, table):
	src_folders = os.listdir(src)
	total=0
	not_found=0
	copy_failed=0
	nrows = table.nrows
	
	for row in range(nrows):
		found = False
		row_data = table.row_values(row)
		sku = row_data[0]
		type = row_data[1]
		sku = str(int(row_data[0]))
		for name in src_folders:			
			if name.find(sku) > -1:
				copy_jpg_to_dest(src + "\\" + name, sku, type)
				total+=1
				found = True
				break
		
		if not found:
			for name in src_folders:
				if os.path.isdir(src + "\\" + name):
					inner_folders = os.listdir(src + "\\" + name)
					for inner_name in inner_folders:			
						if inner_name.find(sku) > -1:
							succeed = copy_jpg_to_dest(src + "\\" + name + "\\" + inner_name, sku, type)
							if succeed:
								total+=1
								found = True
							else:
								copy_failed+=1
							break			
		if not found:
			log(1, "Not found:", sku)
			not_found+=1
	
	log(1, "JPG copy Total:" , str(total))
	log(1, "JPG copy failed  Total:" , str(copy_failed))
	log(1, "JPG not found  Total:" , str(not_found))
	
def copy_jpg_to_dest(path, sku, type):
	dst = "Model library"
	if type:
		dst = dst + "\\" + type
	if not os.path.exists(dst):
		os.makedirs(dst)
		
	src_folders = os.listdir(path)
	if len(src_folders) == 0:
		log(1, "No jpg found for", sku)
		return False
		
	for jpg in src_folders:
		if jpg.find("0.jpg") > -1:
			log(1, "Copy jpg:", sku)
			shutil.copy(path + "\\" + jpg, dst + "\\" + sku + ".jpg")
			return True
	log(1, "Folder found not jpg for", sku)
	return False
	
	
def copy_zip_to_dest(path, sku, type):
	dst = "Model library"
	if type:
		dst = dst + "\\" + type		
	if not os.path.exists(dst):
		os.makedirs(dst)
		
	src_folders = os.listdir(path)
	if len(src_folders) == 0:
		log(1, "No folder found!", path)
		return False
		
	#the last one is newest	
	newest = src_folders[-1]
	if os.path.isdir(path + "\\" + newest):
		zips = os.listdir(path + "\\" + newest)
		for zip in zips:
			if zip.find(sku) > -1 and zip.find(".zip") > -1:
				shutil.copy(path + "\\" + newest + "\\" + zip, dst + "\\" + sku + ".zip")	
				shutil.copystat(path + "\\" + newest + "\\" + zip, dst + "\\" + sku + ".zip")				
				return True		
	else:
		if newest.find(sku) > -1 and newest.find(".zip") > -1:
			shutil.copy(path + "\\" + newest, dst + "\\" + sku + ".zip")	
			shutil.copystat(path + "\\" + newest, dst + "\\" + sku + ".zip")
			return True	
	return False

def remove_cache_files(src):	
	src_folders = os.listdir(src)
	for name in src_folders:
		if os.path.isdir(src + "\\" + name):
			remove_cache_files(src + "\\" + name)
		else:
			if name == "Thumbs.db" or name == "Thumbs_db":
				os.remove(src + "\\" + name)

def log(debug=0, *args):
	if debug == 0:
		pass
		# print(*args)
	else:
		print(*args)
		for arg in args:
			fo.write(str(arg))
		fo.write("\r")
	
def usage():
	print('''
here is the script for copy sku zip and jsp to category folder.
1. read sku from "SKU List.xlsx" and the sheet named SKU
	55282930	furniture
	28648021	bedroom & makeup vanities

2. find sku zip file from "Model Deliverables" folder. find the sku zip file by the zip file name. and then copy the newest zip file to "<sku>.zip"

3. find sku jpg file from "imgs" folder. find the sku jpg file by the same folder name and then copy the "0.jpg" to "<sku>.jpg"

4. all zip and jpg copied to "Model library" folder

5. log will be saved to log.txt

example: python ModelDistribution.py
	''')
	
def main():	

	if len(sys.argv) > 1:
		usage()
		return
		
	date = "SKU"
		
	path = os.getcwd()
		
	log(1, "Remove all cached file Thumbs.db...")
	remove_cache_files(path+"\\Model Deliverables")
	
	log(1, "Read sku from xls...")
	raw_data = open_xls(path + "\\SKU List.xlsx")		
	
	table = read_sheet(raw_data, date)
	
	find_sku_zip(path, table)
	
	log(1, "")
	log(1, "====================================================")
	log(1, "")
	
	find_sku_jpg(path+"\\imgs", table)
	
if __name__ == "__main__":
	main()	