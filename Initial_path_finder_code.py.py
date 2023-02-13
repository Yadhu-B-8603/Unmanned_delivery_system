from PIL import Image
from pytesseract import pytesseract
import time
import pandas as pd
from openpyxl import load_workbook

#Define path to tessaract.exe
path_to_tesseract = r"C:\\Users\\ASUS\\AppData\\Local\\Tesseract-OCR\\tesseract.exe"

#Define path to image
path_to_image = "C:\\Users\\ASUS\\Downloads\\WhatsApp Image 2022-10-15 at 07.29.42.jpeg"

#Point tessaract_cmd to tessaract.exe
pytesseract.tesseract_cmd = path_to_tesseract

hostel_name = ['Zircon','Beryl','Pearl','Opal']
hostel_path = {'Zircon':[1,5,9,13,14],'Beryl':[1,2],'Pearl':[1,5,9,10,11],'Opal':[1,17]}

#Open image with PIL
img = Image.open(path_to_image)

#Extract text from image
text = pytesseract.image_to_string(img)


for string in text.split(" "):
    if string in hostel_name:
        break

time_now = time.ctime(time.time())
date_list = []
for t in time_now.split(" "):
    date_list.append(t)
database_for_tracking = pd.ExcelWriter("C:\\Users\\ASUS\\OneDrive\\Desktop\\Database_Order_tracking.xlsx",engine = 'xlsxwriter')
df = pd.DataFrame({'Name': [string],
                   'Time':[date_list[3]],
                   'Date':[str(date_list[0]+" "+date_list[1]+" "+date_list[2]+" "+date_list[4])],
                   'Status':["Out for delivery"]})
writer = pd.ExcelWriter('C:\\Users\\ASUS\\OneDrive\\Desktop\\Database_Order_tracking.xlsx', engine='openpyxl')
writer.book = load_workbook('C:\\Users\\ASUS\\OneDrive\\Desktop\\Database_Order_tracking.xlsx')
writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
reader = pd.read_excel('C:\\Users\\ASUS\\OneDrive\\Desktop\\Database_Order_tracking.xlsx')

df.to_excel(writer,index=False,header=False,startrow=len(reader)+1)

writer.close()
