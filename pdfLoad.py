import os
import tabula
import pandas as pd

path = os.getcwd() + '\\Reports'
filename = os.listdir(path)[0]
filepath = path + '\\' + filename
excelpath = path + '\\' + filename[:-3] + 'xlsx'

def pdf_table_to_xlsx(pdf_path, xlsx_path, export=True):

    tables = tabula.read_pdf(pdf_path, pages='all')
    if export == True:

        writer = pd.ExcelWriter(xlsx_path, engine='xlsxwriter')
        table_index = 1
        for table in tables:
            DF = pd.DataFrame(table)
            DF.to_excel(writer, sheet_name = 'Sheet{}'.format(table_index))
            table_index = table_index+1
        writer.save()
    else:
        return tables

if __name__ == '__main__':
    path = os.getcwd() + '\\Reports'
    filename = os.listdir(path)[0]
    filepath = path + '\\' + filename
    excelpath = path + '\\' + filename[:-3] + 'xlsx'
    pdf_table_to_xlsx(filepath, excelpath)

