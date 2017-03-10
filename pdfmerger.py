from PyPDF2 import PdfFileMerger

pdfs = ['app3.pdf', 'app333.pdf']

merger = PdfFileMerger()

for pdf in pdfs:
    merger.append(open(pdf, 'rb'))

with open('result.pdf', 'wb') as fout:
    merger.write(fout)
