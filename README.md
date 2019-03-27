## PythonTool

Python tool is a python tools set.
Each file is a tool to help user finish one task.


##### ModelDistribution.py
here is the script for copy sku zip and jpg to category folder.
1. read sku from "SKU List.xlsx" and the sheet named SKU
	55282930	furniture
	28648021	bedroom & makeup vanities
2. find sku zip file from "Model Deliverables" folder. find the sku zip file by the zip file name. and then copy the newest zip file to "<sku>.zip"
3. find sku jpg file from "imgs" folder. find the sku jpg file by the same folder name and then copy the "0.jpg" to "<sku>.jpg"
4. all zip and jpg copied to "Model library" folder
5. log will be saved to log.txt
example: python ModelDistribution.py
	''')


##### TeamTaskAssignment.py

there are many models need to do. the models have different level. they had been split into 10 levels. this tool to help us assgin models to different pepole with save level models.


##### issues2googledrive*.py

this tool help us find model zip file from excel and move this zip file to another place.
