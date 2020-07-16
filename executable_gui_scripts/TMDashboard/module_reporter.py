import PyPDF2
import re
import os

report_typelist = ['bepcb', 'cv', 'lc']


def getPageList(pdf, keyword):

    pagecount = pdf.getNumPages()
    print("Number of pages: {}".format(pagecount))
    pagelist = []
    for i in range(0, pagecount):
        page = pdf.getPage(i)
        text = page.extractText()
        ResSearch = re.search(keyword, text)
        if(ResSearch != None):
            # insert page num of the first page of each report into list
            pagelist.append(i)
    return pagelist, pagecount


def getStaffID(page, keyword):
    if keyword == 'Assessee ID:':
        nextWord = 'Position'
    if keyword == 'STAFF NO.:':
        nextWord = 'POSITION'

    keywordLength = len(keyword)
    pagetext = page.extractText()

    keyword_startindex = pagetext.find(keyword)
    nextword_startindex = pagetext.find(nextWord)

    start = keyword_startindex + keywordLength
    end = nextword_startindex
    staffID = pagetext[start:end].lstrip('0').lstrip()
    return staffID


def splitPdf(pdf, pagelist, pagecount, reportpath, keyword, mainpath):
    outputfolder = os.path.join(mainpath, reportpath)
    if not os.path.exists(outputfolder):
        os.makedirs(outputfolder)
    outputlist = []
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
        staffID = getStaffID(firstPage, keyword)
        # print(staffID)
        filename = "%s.pdf" % staffID
        filepath = os.path.join(outputfolder, filename)
        with open(filepath, "wb") as outputStream:
            output.write(outputStream)
        outputlist.append(filepath)
    print("DONE")
    return outputlist, outputfolder


def mainProcess(inputfile, selection_index, mainpath):

    # report_typelist = ['bepcb', 'cv', 'lc']
    keywordlist = ['Assessee ID:', 'STAFF NO.:', 'Assessee ID:']
    report_type = report_typelist[selection_index]
    keyword = keywordlist[selection_index]

    object = PyPDF2.PdfFileReader(inputfile)
    pagelist, pagecount = getPageList(object, keyword)
    # print(pagelist)
    outputlist, outputfolder = splitPdf(
        object, pagelist, pagecount, report_type, keyword, mainpath)
    return outputlist, outputfolder


# inputfile = "bepcb21.pdf"
# inputfile = r'D:\Documents\Python\project-2\database\input\bepcb21.pdf'
# selection_index = 0
# outputlist = mainProcess(inputfile, selection_index)
# mainpath = os.getcwd()
# input_path = os.path.join(mainpath, "database/input/", inputfile)
# outputdir = os.path.join(mainpath, "database/output/")
# if not os.path.exists(outputdir):
#     os.makedirs(outputdir)
# object = PyPDF2.PdfFileReader("bepcb21.pdf")
# inputpdf = PyPDF2.PdfFileReader(open("bepcb21.pdf", "rb"))

# object = PyPDF2.PdfFileReader(inputfile)
# pagelist, pagecount = getPageList(object)
# print("number of report in pdf: {}".format(len(pagelist)))
# print(pagelist)
# print(pagecount)
# splitPdf(object, pagelist, pagecount)
