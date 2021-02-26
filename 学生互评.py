import pandas as pd
import xlwt


grade_1 = pd.read_excel(r"D:\综合素质评价\2020.12月综合素质评价\学生互评\高三.xls")
#stuNum = [0,66,66,65,66,65,64,65,64,65,63,63,62,64,63,63,63,61,62,62,59,56,57,59] #23届
#stuNum = [0,65,65,65,65,65,64,64,64,64,64,62,61,65,45,47,50,50,62,64,21]  #22届
stuNum = [0,62,61,66,60,63,64,57,53,54,53,54,62,49,57,59,56,65,65,48,68,17,9]  #21届
grade_end = []
maxnum=24
for (class_, group) in grade_1.groupby('班级'):
    if int(class_)<maxnum:
        #print(class_)
        grade_end.extend(((group[group.columns[2:stuNum[int(class_)] + 2]].mean())*10).values)
f = xlwt.Workbook()
sheet1 = f.add_sheet(u'sheet111', cell_overwrite_ok=True)  # 创建sheet
# 将数据写入第 i 行，第 j 列
i=1
for grade in grade_end:
    sheet1.write(i,2,grade )
    i = i + 1
f.save(r"D:\综合素质评价\2020.12月综合素质评价\学生互评\高三.xls")  # 保存文件
print("完成")
