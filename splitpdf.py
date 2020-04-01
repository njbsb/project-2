import PyPDF2
import re
import os
from pathlib import Path


def getStaffID(page):
    assID = "Assessee ID:"
    assID_len = len(assID)
    text = page.extractText()
    asid_id = text.find(assID)
    start = asid_id + assID_len
    end = start + 8
    staffID = text[start:end].lstrip("0")
    return staffID


def getPageList(pdf):
    pagecount = pdf.getNumPages()
    print("Number of pages: {}".format(pagecount))
    String = "STAFF DETAILS"
    pagelist = []
    for i in range(0, pagecount):
        page = pdf.getPage(i)
        text = page.extractText()
        ResSearch = re.search(String, text)
        if(ResSearch != None):
            # insert page num of the first page of each report into list
            pagelist.append(i)
            # print(ResSearch)
    return pagelist, pagecount


inputfile = "bepcb21.pdf"
mainpath = os.getcwd()
input_path = os.path.join(mainpath, "database/input/", inputfile)
# object = PyPDF2.PdfFileReader("bepcb21.pdf")
object = PyPDF2.PdfFileReader(input_path)
# NumPages = object.getNumPages()
pagelist, pagecount = getPageList(object)

print("number of report in pdf: {}".format(len(pagelist)))
print(pagelist)

# inputpdf = PyPDF2.PdfFileReader(open("bepcb21.pdf", "rb"))
# pagecount = PyPDF2.PdfFileReader("bepcb21.pdf").getNumPages()
# pagecount2 = inputpdf.getNumPages()
outputdir = os.path.join(mainpath, "database/output/")
if not os.path.exists(outputdir):
    os.makedirs(outputdir)

for i in range(len(pagelist)):
    output = PyPDF2.PdfFileWriter()
    if(i < len(pagelist)-1):  # check if i is the last index
        # diff is no of page for each sliced pdf
        diff = pagelist[i+1] - pagelist[i]
    else:
        diff = pagecount - pagelist[i]
    for j in range(diff):
        k = pagelist[i] + j
        output.addPage(object.getPage(k))

    # get id/page
    pageObj = object.getPage(pagelist[i])
    staffID = getStaffID(pageObj)
    # print(staffID)
    filename = "%s.pdf" % staffID
    with open(outputdir + filename, "wb") as outputStream:
        output.write(outputStream)
print("DONE")
