import os
from pdf2image import convert_from_path
import cv2
import pytesseract
import re
import pandas as pd
import requests
from urllib.parse import urlparse

# pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

pdfDirPath = "Pdf"
imageDirPath = "Image"
excelDirPath = "Excel"

excelFilePath = os.path.join(excelDirPath, "Sample.xlsx")
newExcelFilePath = os.path.join(excelDirPath, "Sample2.xlsx")
print("READING EXCEL {0}...".format(excelFilePath))
df = pd.read_excel(excelFilePath)

print("REMOVING CURRENT PDF FILES")
pdfFiles = os.listdir(pdfDirPath)
for pdfFile in pdfFiles:
	os.remove(os.path.join(pdfDirPath, pdfFile))

print("REMOVING CURRENT IMAGE FILES")
imageFiles = os.listdir(imageDirPath)
for imageFile in imageFiles:
	os.remove(os.path.join(imageDirPath, imageFile))
	
if os.path.exists(newExcelFilePath):
	os.remove(newExcelFilePath)

newDF = {}
newDF['Link'] = []
newDF['FEI Number'] = []
for i in df.index:
	try:
		urlLink = df['Link'][i]
		urlParse = urlparse(urlLink)
		baseName = os.path.basename(urlParse.path)
		name = baseName.rstrip('.pdf')
		imgFilePath = os.path.join(imageDirPath, name + ".jpg")
		pdfFilePath = os.path.join(pdfDirPath, baseName)
		
		print("SENDING REQUEST: {0}".format(urlLink))
		response = requests.get(urlLink, stream=True)
		print("WRITING PDF: {0}...".format(pdfFilePath))
		with open(pdfFilePath, 'wb') as f:
			f.write(response.content)
			
		print("CONVERTING TO IMAGE {0}...".format(imgFilePath))
		images = convert_from_path(pdfFilePath,last_page=1,dpi=300,fmt="jpeg")
		image1 = images[0]
		image1.save(imgFilePath, 'JPEG') 
		
		print("SCANNING IMAGE...")
		img = cv2.imread(imgFilePath, cv2.IMREAD_GRAYSCALE)
		print("FINDING TEXT WITH TEN DIGITS...")
		everyOcc = re.findall("\d{10}", pytesseract.image_to_string(img, config="--psm 6"))
		
		newDF['Link'].append(urlLink)
		if len(everyOcc) > 0:
			print("TEXT FOUND - ADDING TO NEW DF")
			newDF['FEI Number'].append(everyOcc[0])
		else:
			print("TEXT NOT FOUND")
			newDF['FEI Number'].append("")
	except Exception as e:
		print(e)

print("WRITING EXCEL {0}...".format(newExcelFilePath))
df = pd.DataFrame(newDF)
df.to_excel(newExcelFilePath)

