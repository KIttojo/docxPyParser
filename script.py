import os
# import zipfile
import xmltodict, json

import shutil
import sys



try:
	os.rename('test.docx', 'test.zip')

	with open(".\\targetdir\\word\\document.xml", encoding='utf-8') as file:
		XML_content = file.read()
		XML_body_indx = XML_content.find("<w:body>")
		XML_content = XML_content[XML_body_indx:].replace("\n", "").removesuffix("</w:document>")
		print(XML_content)

		o = xmltodict.parse(XML_content)
		res = json.dumps(o)
		print(res)

		res1 = open("res.txt", "w")
		res1.write(res)
		res1.close()


	print("OK")

except:
	print("Error occured while rurring script")