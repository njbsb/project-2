import PyPDF2
import re

# open the pdf file
object = PyPDF2.PdfFileReader("bepcb1.pdf")

# get number of pages
NumPages = object.getNumPages()
print(NumPages)

String = "STAFF DETAILS"
pagelist = []
# extract text and do the search
for i in range(0, 20):
    PageObj = object.getPage(i)
    Text = PageObj.extractText()
    ResSearch = re.search(String, Text)
    if(ResSearch != None):
        pagelist.append(i)
        # print(ResSearch)
print("number of report in pdf: {}".format(len(pagelist)))
print(pagelist)

inputpdf = PyPDF2.PdfFileReader(open("bepcb21.pdf", "rb"))
pagecount = PyPDF2.PdfFileReader("bepcb21.pdf").getNumPages()
for i in range(len(pagelist)):
    output = PyPDF2.PdfFileWriter()
    # check if i is the last index
    if(i < len(pagelist)-1):
        diff = pagelist[i+1] - pagelist[i]
        print(diff)
        for j in range(diff):
            k = pagelist[i] + j
            output.addPage(inputpdf.getPage(k))
    else:
        diff = pagecount - pagelist[i]
        print(diff)
        for j in range(diff):
            k = pagelist[i]+j
            output.addPage(inputpdf.getPage(k))
    with open("document-page%s.pdf" % i, "wb") as outputStream:
        output.write(outputStream)
