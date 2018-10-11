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
    
def find_sku_zip(path, table, date):	
    src = path+"\\Model Deliverables"
    total=0
    not_found=0
    copy_failed=0
    nrows = table.nrows
    
    for row in range(nrows):
        found = False
        row_data = table.row_values(row)
        sku = row_data[0]
        sku = str(int(row_data[0]))
        zip_path = find_one_sku_zip(src, sku)
        if zip_path != None:
            copied = copy_zip_to_dest(zip_path, sku, date)
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
    
    
def copy_zip_to_dest(path, sku, date):
    dst = "Deliver file\\" + date
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
                fromPath = path + "\\" + newest
                toPath = get_new_path(fromPath, dst)
                log(1, "move:", fromPath, toPath)
                if not os.path.exists(toPath):
                    os.makedirs(toPath)
                shutil.move(fromPath + "\\" + zip, toPath + "\\" + zip)	
                #shutil.copystat(path + "\\" + newest + "\\" + zip, dst + "\\"  + zip)
                move_retexturable(path + "\\" + newest, dst)
                remove_folder(fromPath, path)
                return True		
    else:
        if newest.find(sku) > -1 and newest.find(".zip") > -1:
            fromPath = path
            toPath = get_new_path(fromPath + "\\" + newest, dst + "\\" + newest)
            log(1, "move without time folder:", fromPath, toPath)
            if not os.path.exists(toPath):
                os.makedirs(toPath)
            shutil.move(fromPath, toPath)	
            #shutil.copystat(path + "\\" + newest, dst + "\\" + zip)
            move_retexturable(path + "\\" + newest, dst)
            remove_folder(fromPath, path)
            return True	
    return False

def get_new_path(fromPath, new):
    old = "Model Deliverables"
    toPath = fromPath.replace(old, new)
    return toPath
    
def remove_folder(fromPath, path):
    src_folders = os.listdir(fromPath)
    if len(src_folders) == 0:
        log(1, "remove folder:", path)
        shutil.rmtree(path)
    
def move_retexturable(path, dst):
    src_folders = os.listdir(path)
    if len(src_folders) == 1:
        file = src_folders[0]
        if file.lower().find("rooomy.zip") > -1:
            fromPath = path + "\\" + file
            toPath = get_new_path(fromPath, dst)
            log(1, "move retexturable:", fromPath, toPath)
            shutil.move(fromPath, toPath)	
    elif len(src_folders) > 1:
        for file in src_folders:
            if file.lower().find("rooomy.zip") > -1:
                fromPath = path + "\\" + file
                toPath = get_new_path(fromPath, dst)
                log(1, "copy retexturable:", fromPath, toPath)
                shutil.copy(fromPath, toPath)
     

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

    date = sys.argv[1]
        
    path = os.getcwd()
        
    log(1, "Remove all cached file Thumbs.db...")
    remove_cache_files(path+"\\Model deliverables")
    
    log(1, "Read sku from xls...")
    raw_data = open_xls(path + "\\QACustomer.xls")			
    
    table = read_sheet(raw_data, date)
    
    find_sku_zip(path, table, date)
    
if __name__ == "__main__":
    main()	
    