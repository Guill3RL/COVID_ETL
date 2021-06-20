import os
import tabula
import pandas as pd

path = os.getcwd() + '\\Reports'
filename = os.listdir(path)[0]
filepath = path + '\\' + filename
excelpath = path + '\\' + filename[:-3] + 'xlsx'

def pdf_table_to_xlsx(pdf_path, xlsx_path=None, export=True):

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

def scan_files(path, pattern):
    pdf_list = [pdf for pdf in os.listdir(path) if pattern in pdf]

    pdf_scan= []

    for pdf in pdf_list:
        tables = pdf_table_to_xlsx(path + '\\' + pdf, xlsx_path=None, export=False)
        pdf_characteristics = {'Name':pdf}
        pdf_characteristics['Total Tables'] = len(tables)

        tables_characteristics = []
        for table in tables:
            tables_characteristics.append(table.size)

        pdf_characteristics['Tables Sizes'] = tables_characteristics
        pdf_scan.append(pdf_characteristics)
        print(pdf + ' Completed!')
    return pdf_scan

if __name__ == '__main__':
    path = os.getcwd() + '\\COVID'
    pattern = 'Actualizacion_10'
    pdf_characteristics = scan_files(path, pattern)
    print(pdf_characteristics)
    #pdf_list = [pdf for pdf in os.listdir(path) if pattern in pdf]
    #tables = pdf_table_to_xlsx(path + '\\' + pdf_list[0], xlsx_path=None, export=False)
    #print(len(tables))

