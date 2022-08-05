import openpyxl
import os

filelist = []
sheet_name = 'Test'
for root, dirs, files in os.walk(".", topdown=False):
    for name in files:
        str = os.path.join(root, name)
        if str.split('.')[-1] == 'xlsx':
            filelist.append(str)

for i in range(len(filelist)):
    workbook = openpyxl.load_workbook(filelist[i])
    # 删除目标Sheet
    if sheet_name in workbook:
        worksheet = workbook[sheet_name]
        workbook.save(filelist[i])
        workbook.remove(worksheet)
        print(filelist[i]+' delete successfully!')
    else:
        print(filelist[i]+'的指定Sheet不存在，故不作处理')
