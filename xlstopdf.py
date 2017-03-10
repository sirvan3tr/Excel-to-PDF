## Author: Sirvan Almasi Jan 2017
## This script helps in automating the process of converting an excel into PDF
import win32com.client, time

o = win32com.client.Dispatch("Excel.Application")
o.Visible = False
timedate = time.strftime("%H%M__%d_%m_%Y")
wb_path = r'S:/GSA Euro Research Company Files/Property Sectors/Euro Office Sector/London Offices/Green Street Research Reports/London Office Report Feb 17/_5. Appendix - Company Snapshots - Copy.xlsm'
#wb_path = r'C:/Users/salmasi/Documents/MATLAB/xlstopdf/22.xlsm'
wb = o.Workbooks.Open(wb_path)

ws_index_list = [1,2,3] #say you want to print these sheets
path_to_pdf = r'C:/Users/salmasi/Documents/MATLAB/xlstopdf/app__'+str(timedate)+'.pdf'
wb.WorkSheets(ws_index_list).Select()
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
wb.Close(True)
