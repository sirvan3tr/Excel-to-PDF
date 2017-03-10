## Author: Sirvan Almasi Jan 2017
## This script helps in automating the process of converting an excel into PDF
import win32com.client, time, os, csv, random
from PyPDF2 import PdfFileMerger

print('Enter the directory of your csv')
files_to_pdf = input()
#files_to_pdf = 'files_to_pdf.csv'
pdfs = []

with open(files_to_pdf, 'r') as f:
    reader = csv.reader(f, delimiter=',')
    count = 0
    for excel in reader:
        ## ----
        if count == 0:
            timedate = time.strftime("%d-%m-%Y %H.%M")
            pub_path = excel[0]+timedate
            if not os.path.exists(pub_path):
                os.makedirs(pub_path)
        ## ----
        o = win32com.client.Dispatch("Excel.Application")
        o.Visible = False
        #random number for pdf name
        random_int = random.randint(0,1000)

        wb_path = str(excel[0])+str(excel[1])
        pdfs.append(pub_path + '/'+str(excel[1])+str(random_int)+'.pdf') # for later to combine pdfs
        wb = o.Workbooks.Open(wb_path)

        #ws_index_list = [1,2,3] #say you want to print these sheets
        ## split page numbers
        ws_index_list = []
        page_num = excel[2].split(',')
        for num in page_num:
            ws_index_list.append(int(num))
        ## --- end
        path_to_pdf = pub_path + '/'+str(excel[1])+str(random_int)+'.pdf'
        wb.WorkSheets(ws_index_list).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
        wb.Close(True)
        count = count + 1


merger = PdfFileMerger()
for pdf in pdfs:
    merger.append(open(pdf, 'rb'))

with open(pub_path+'/Final.pdf', 'wb') as fout:
    merger.write(fout)
