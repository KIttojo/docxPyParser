import os
import xmltodict, json
import zipfile

try:
	os.rename('test2.docx', 'test.zip')

	with zipfile.ZipFile(".\\test.zip", 'r') as zip_ref:
		zip_ref.extractall(".\\targetdir")
		with open(".\\targetdir\\word\\document.xml", encoding='utf-8') as file:
			XML_content = file.read()
			XML_body_indx = XML_content.find("<w:body>")
			XML_content = XML_content[XML_body_indx:].replace("\n", "").removesuffix("</w:document>")

			o = xmltodict.parse(XML_content)
			res = json.dumps(o, ensure_ascii=False)
			# print(res)
			body = json.loads(res)

			content = body["w:body"]["w:p"]
			
			lines = []

			for item in content:
				ttype = ""

				for tag, val in item.items():
					if tag == "w:pPr":
						ttype = "Title"
						print("!!")
						lines.append(ttype)
						break
						# isparagraph = True;
						# val2 = str(val).replace("'", '"')
						# print(val2) #val - dict под видом объекта, а должен быть строкой
						# temp = json.loads(str(val2))
						# print(type(temp))
						# text = temp["w:r"]["w:t"]["w:pStyle"]["@w:val"]
						# print(text)
					else:
						ttype = "Paragraph"
						lines.append(ttype)
				# lines.append(ttype)
			
			print(lines)

			res1 = open("res.txt", "w")
			res1.write(body)
			res1.close()

	os.rename('test.zip', 'test2.docx')

	print("OK")

except Exception as inst:
	os.rename('test.zip', 'test2.docx')
	print("Error occured while rurring script", inst)