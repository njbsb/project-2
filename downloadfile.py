import urllib.request

id = [1054700, 1047239, 1046975, 1046260]

for i in id:
    filename = str(i) + ".pdf"
    filelink = "https://talentengine.petronas.com/api/resources/staff/" + \
        str(i) + "/cv"
    urllib.request.urlretrieve(
        filelink, filename)


# filename = str(id) + ".pdf"
# filelink = "https://talentengine.petronas.com/api/resources/staff/" + \
#     str(id) + "/cv"
# print(filelink)
# urllib.request.urlretrieve(
#     filelink, filename)
