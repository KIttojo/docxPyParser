import os
import zipfile

try:
	os.rename('test.docx', 'test.zip')

	with zipfile.ZipFile("test.zip", "r") as zip_ref:
		zip_ref.extractall("targetdir")
	print("OK")

except:
	print("Error occured while rurring script")