import urllib.request

id = [105363, 100296, 103390, 121862, 114164, 122791, 119607, 116893, 104567, 104827, 115405, 100886, 107732, 133834, 116090, 126205, 124622, 127157, 125131
      ]

for i in id:
    filename = str(i) + ".jpg"
    filelink = "https://talentengine.petronas.com/api/resources/staff/" + \
        str(i) + "/picture"
    urllib.request.urlretrieve(
        filelink, filename)


# filename = str(id) + ".pdf"
# filelink = "https://talentengine.petronas.com/api/resources/staff/" + \
#     str(id) + "/cv"
# print(filelink)
# urllib.request.urlretrieve(
#     filelink, filename)
