import os
import pikepdf as pk
import tabula
import pandas as pd

path = os.getcwd() + '\\Reports'
filename = os.listdir(path)[0]
filepath = path + '\\' + filename
excelpath = path + '\\' + filename[:-3] + 'xlsx'

tables = tabula.read_pdf(filepath, pages='all')

writer = pd.ExcelWriter(excelpath, engine='xlsxwriter')
table_index = 1
for table in tables:
    DF = pd.DataFrame(table)
    DF.to_excel(writer, sheet_name = 'Sheet{}'.format(table_index))
    table_index = table_index+1
writer.save()

#print(excelpath)
#file = pk.Pdf.open(filepath)
#DF.to_excel('table1.xlsx')
#pagenumber = 1
#for page in file.pages:
#page = file.pages[0]
#print(repr(page))
    #os.mkdir(path + '/page{}'.format(pagenumber))
    #os.chdir (path + '/page{}'.format(pagenumber))
    #for image in page.images.keys():
    #    rawimage = page.images[image]
    #    pdfimage = pk.PdfImage(rawimage)
    #    pdfimage.extract_to(fileprefix=image[1:])
    #pagenumber = pagenumber + 1


