from PIL import Image
from PyPDF2 import PdfFileMerger, PdfFileReader

for i in range(1,301):
    image1 = Image.open(r"C:/Users/inci1/Desktop/t60w/bcs_"+str(i)+".png")
    im1 = image1.convert('RGB')
    im1.save(r"C:/Users/inci1/Desktop/pdfs/bcs_"+str(i)+".pdf")

    
mergedObject = PdfFileMerger()

for i in range(1,301):
    mergedObject.append(PdfFileReader("C:/Users/inci1/Desktop/pdfs/bcs_"+str(i)+".pdf", 'rb'))
 
mergedObject.write("C:/Users/inci1/Desktop/t60w.pdf")