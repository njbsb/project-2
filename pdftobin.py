import base64 as bs

def convert_image_tobase64(imgname, base64name):
    with open(imgname, "rb") as pdf_file:
        encoded_string = bs.b64encode(pdf_file.read())

    f = open(base64name, 'wb')
    f.write(encoded_string)
    f.close()
    return encoded_string

def bintoimg(encod, newname):
    bytes = bs.b64decode(encod, validate=True)
    f = open(newname, 'wb')
    f.write(bytes)
    f.close()

def convert_pdf_tobytestring(filename, bytetextname):
    pdfname = filename+ ".pdf"
    with open(pdfname, "rb") as pdf_file:
        encoded_string = bs.b64encode(pdf_file.read())
    
    txtname = bytetextname + ".txt"

    f = open(txtname, 'wb')
    f.write(encoded_string)
    f.close()
    return encoded_string

def convert_bytestring_topdf(encodedstring, filenewname):
    bytes = bs.b64decode(encodedstring, validate=True)
    if bytes[0:4] != b'%PDF':
        raise ValueError('Missing the PDF file signature')

    filename = filenewname + ".pdf"
    f = open(filename, 'wb')
    f.write(bytes)
    f.close()

# s = convert_image_tobase64("image.jpg", "imagebyte.txt")
# print(s)

# str_byte = convert_pdf_tobytestring("book", "byte")
# convert_bytestring_topdf(str_byte, "newpdf")