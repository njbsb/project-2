import win32com.client as win32

def convert_to_xlsx(filename):
    input_path = r'C:\Users\mnajib.suib\Documents\Python\project-2\database\input\\'
    fname = input_path[:-1] + filename
    print(fname)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()

def convert_to_xls(filename):
    input_path = r"C:\Users\mnajib.suib\Documents\Python\project-2\database\input\\"
    fname = input_path[:-1] + filename
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(fname[:-1], FileFormat = 56)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()

'''
fname = ""
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)
wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()
'''