from os import name
from re import template
from typing import Sized
import pandas as pd
import code128
from numpy import column_stack, nan
from PIL import Image, ImageFont, ImageDraw 
from PyPDF2 import PdfFileMerger, PdfFileReader
import email_listener
from time import sleep
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from zipfile import ZipFile

sleep(2)

email = "barcodelistener@gmail.com"
app_password = "50773fe6"
folder = "Inbox"
attachment_dir = "C:/Users/inci1/Desktop"
el = email_listener.EmailListener(email, app_password, folder, attachment_dir)

el.login()

messages = el.scrape()
print(messages)

sleep(2)

data = pd.read_excel (r'C:/Users/inci1/Desktop/liste.xlsx',dtype=str,convert_float=1,index_col=None,header=None)

imei = data[0].tolist()
imei = [item for item in imei if not(pd.isnull(item)) == True]

names = data[1].tolist()
names = [item for item in names if not(pd.isnull(item)) == True]

for i in range(len(imei)):


    barkod = code128.image(imei[i])
    bw, bh = barkod.size
    barkod.crop((27, 0, bw-27, bh))

    if "MI" in names[i]:
        template1= Image.open('C:/Users/inci1/Documents/barkcodemaker/kodlu.jpg')
    else:
        template1= Image.open('C:/Users/inci1/Documents/barkcodemaker/kodsuz.jpg')

    tw,th = template1.size

    title_font = ImageFont.truetype('C:/Users/inci1/Documents/barkcodemaker/regular.ttf', 30)
    imei_font = ImageFont.truetype('C:/Users/inci1/Documents/barkcodemaker/regular.ttf', 30)

    template1.paste(barkod,(480, 150))
    template1.paste(barkod,(480, 95))

    image_editable = ImageDraw.Draw(template1)

    splitted = names[i].split(" ")
    for j in range(len(splitted)):
        if("GB" in splitted[j]):
            devam=names[i].split(splitted[j])
            devam[1]= (splitted[j]+devam[1])
            #print(devam[0])
        
    image_editable.text((32,30), devam[0], (0), font=title_font,stroke_width=1,)
    image_editable.text((508,30), devam[1].replace("GB"," GB"), (0), font=title_font,stroke_width=1)
    image_editable.text((508,263), ("IMEI:          "+imei[i]), (0), font=imei_font,stroke_width=1,)


    template1.save("C:/Users/inci1/Documents/barkcodemaker/pdfs/bcs_"+str(i)+".pdf")

mergedObject = PdfFileMerger()

for i in range(len(imei)):
    mergedObject.append(PdfFileReader('C:/Users/inci1/Documents/barkcodemaker/pdfs/bcs_' + str(i)+ '.pdf', 'rb'))
 
mergedObject.write("C:/Users/inci1/Desktop/liste.pdf")

'''
sleep(2)

body = "bak bakim olmuş mu"
sender = 'barcodelistener@gmail.com'
password = '50773fe6'
receiver = 'inci1memet@gmail.com'

message = MIMEMultipart()
message['From'] = sender
message['To'] = receiver
message['Subject'] = "mail'i ben göndermedim bot gönderdi"

message.attach(MIMEText(body, 'plain'))

pdfname = 'C:/Users/inci1/Desktop/liste.pdf'

binary_pdf = open(pdfname, 'rb')

payload = MIMEBase('application', 'octate-stream', Name=pdfname)
payload.set_payload((binary_pdf).read())

encoders.encode_base64(payload)

payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
message.attach(payload)

session = smtplib.SMTP('smtp.gmail.com', 587)

session.starttls()

session.login(sender, password)

try:
    text = message.as_string()
    session.sendmail(sender, receiver, text)

except:
    print("dosya büyük yapamadım")

session.quit()
print('Mail Sent')
'''