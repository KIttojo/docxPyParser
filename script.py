import os
import xmltodict, json
import zipfile

try:
	inifialfilename = 'test_old.docx'
	tempzip = 'test.zip'

	os.rename(inifialfilename, tempzip)

	with zipfile.ZipFile(f".\\{tempzip}", 'r') as zip_ref:
		zip_ref.extractall(".\\targetdir")
		with open(".\\targetdir\\word\\document.xml", encoding='utf-8') as file:
			XML_content = file.read()
			XML_body_indx = XML_content.find("<w:body>")
			XML_content = XML_content[XML_body_indx:].replace("\n", "").removesuffix("</w:document>")

			o = xmltodict.parse(XML_content)
			res = json.dumps(o, ensure_ascii=False)
			body = json.loads(res)

			content = body["w:body"]["w:p"]
			# print(content)
			
			lines = []
			iobjs = {
				"inf_objects": []
			}

			for item in content:
				# print(item, "\n--------------------------------")
				infobject = {
					"objtype": "",
					"content": "",
				}

				istitle = False
				
				for tag, val in item.items():
					# print("val= ", val)
					#for title
					if tag=="w:pPr":
						try:
							level = val["w:pStyle"]["@w:val"]
							infobject["objtype"] = "title"
							istitle = True
						#if text is "title" but with no style
						except:
							pass
						# infobject["objtype"] = "title"
						# istitle = True

					#for paragraph
					elif tag=="w:r":
						if istitle:
							infobject["content"] = str(val["w:t"])
						else:
							temp = []
							for i in val:
								# print("i= ", i)
								if (i == "w:t"):
									if isinstance(i, dict):
										try: 
											temp.append(i["w:t"]["#text"])
											# print("====", i["w:t"]["#text"])
										except: temp.append(i["w:t"])
									else:
										# print("val = ", val["w:t"])
										temp.append(val["w:t"])
								else:
									try:
										temp.append(i["w:t"]["#text"])
									except:
										temp.append(i["w:t"])

								# print("i= ", i)
								# try: temp.append(i["w:t"]["#text"])
								# except: temp.append(i["w:t"])
							infobject["content"] = ' '.join(temp)
							infobject["objtype"] = "paragraph"
						istitle = False

				# lines.append(infobject)238
				iobjs["inf_objects"].append(infobject)
			# print("--------------------------------\n", iobjs)

			with open('res.json', 'w', encoding='utf-8') as f:
				json.dump(iobjs, f, ensure_ascii=False, indent=4)


	os.rename(tempzip, inifialfilename)


except Exception as inst:
	os.rename(tempzip, inifialfilename)
	print("Error: ", inst.__doc__)
	print(inst.message)