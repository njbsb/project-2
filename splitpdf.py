import PyPDF2
import re
import os

inputfile = "bepcb21.pdf"
mainpath = os.getcwd()
input_path = os.path.join(mainpath, "database/input/", inputfile)
# object = PyPDF2.PdfFileReader("bepcb21.pdf")
object = PyPDF2.PdfFileReader(input_path)

NumPages = object.getNumPages()
print("Number of pages: {}".format(NumPages))
String = "STAFF DETAILS"
pagelist = []
# extract text and do the search
for i in range(0, NumPages):
    PageObj = object.getPage(i)
    Text = PageObj.extractText()
    ResSearch = re.search(String, Text)
    if(ResSearch != None):
        pagelist.append(i)
        # print(ResSearch)
print("number of report in pdf: {}".format(len(pagelist)))
print(pagelist)

# inputpdf = PyPDF2.PdfFileReader(open("bepcb21.pdf", "rb"))
# pagecount = PyPDF2.PdfFileReader("bepcb21.pdf").getNumPages()
# pagecount2 = inputpdf.getNumPages()
assID = "Assessee ID:"
assID_len = len(assID)

for i in range(len(pagelist)):
    output = PyPDF2.PdfFileWriter()
    # check if i is the last index
    if(i < len(pagelist)-1):
        diff = pagelist[i+1] - pagelist[i]
    else:
        diff = NumPages - pagelist[i]
    print(diff)
    for j in range(diff):
        k = pagelist[i] + j
        output.addPage(object.getPage(k))

    # get id/page

    pageObj = object.getPage(pagelist[i])
    pageText = pageObj.extractText()

    assID_index = pageText.find(assID)
    start = assID_index + assID_len
    end = start + 8
    staffID = pageText[start:end]
    print(staffID)
    # print("assessee id found at index: ", assID_index)
    # filename = "document-page%s.pdf" % i
    filename = "BePCB_%s.pdf" % staffID
    outputfile = os.path.join(mainpath, "database/output", filename)
    with open(outputfile, "wb") as outputStream:
        output.write(outputStream)
