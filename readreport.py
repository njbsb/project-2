
# import PyPDF2
# import re

# open the pdf file
# object = PyPDF2.PdfFileReader("dcreport.pdf")

# # get number of pages
# NumPages = object.getNumPages()

# # define keyterms
# String = "Summary"

# # extract text and do the search
# for i in range(0, NumPages):
#     PageObj = object.getPage(i)
#     print("this is page " + str(i))
#     Text = PageObj.extractText()
#     # print(Text)
#     ResSearch = re.search(String, Text)
#     print(ResSearch)
import PyPDF2
pdfFileObj = open('dcreport1.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
print(pdfReader.numPages)

pageObj = pdfReader.getPage(0)
t = pageObj.extractText()
print("Is pdf encrypted: ")
print(pdfReader.getDocumentInfo())

# import textract
# text = textract.process("dcreport.pdf")

# print(text)
'''
from tika import parser

raw = parser.from_file('367.pdf')
st = str(raw)
safe_text = st.encode('utf-8', errors='ignore')
safe_text = str(safe_text).replace("\n", "").replace("\\", "")
print('--- safe text ---')
print(safe_text)
'''
# print(raw['content'])
