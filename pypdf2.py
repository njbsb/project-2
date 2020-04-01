# importing required modules
import PyPDF2

# creating a pdf file object
pdfFileObj = open('document-page0.pdf', 'rb')

# creating a pdf reader object
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

# printing number of pages in pdf file
print(pdfReader.numPages)

# creating a page object
pageObj = pdfReader.getPage(0)

text = pageObj.extractText()
# extracting text from page
print(text)

# closing the pdf file object
pdfFileObj.close()
