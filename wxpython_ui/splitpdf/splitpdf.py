import PyPDF2
import re
import os


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


def splitPdf(pdf, pagelist, pagecount):
    for i, startpage in enumerate(pagelist):
        output = PyPDF2.PdfFileWriter()
        if(i < len(pagelist) - 1):
            diff = pagelist[i+1] - startpage
        else:
            diff = pagecount - startpage
        for page in range(diff):
            p = startpage + page
            output.addPage(pdf.getPage(p))
        firstPage = pdf.getPage(startpage)
        staffID = getStaffID(firstPage)
        # print(staffID)
        filename = "%s.pdf" % staffID
        with open(outputdir + filename, "wb") as outputStream:
            output.write(outputStream)
    print("DONE")


inputfile = "bepcb21.pdf"
mainpath = os.getcwd()
input_path = os.path.join(mainpath, "database/input/", inputfile)
outputdir = os.path.join(mainpath, "database/output/")
if not os.path.exists(outputdir):
    os.makedirs(outputdir)
# object = PyPDF2.PdfFileReader("bepcb21.pdf")
# inputpdf = PyPDF2.PdfFileReader(open("bepcb21.pdf", "rb"))

object = PyPDF2.PdfFileReader(input_path)
pagelist, pagecount = getPageList(object)
print("number of report in pdf: {}".format(len(pagelist)))
print(pagelist)
print(pagecount)
splitPdf(object, pagelist, pagecount)
