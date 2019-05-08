'''
    将数组中的数据遍历，并写入excel表格中
    使用xlwt模块创建并保存excel文件
'''

import xlwt
import os

myData = [{"name":"ZSP","nickName":"寒鸣"},
          {"name":"LH","nickName":"一岁"},
          {"name":"GN","nickName":"宁"},
          {"name":"WHR","nickName":"明儿喂马劈柴"},
          {"name":"FC","nickName":"桃花幻梦"},
          {"name":"WHX","nickName":"Heyson."},
          ]

# 创建workbook
workbook = xlwt.Workbook(encoding='utf-8')
# 创建表
worksheet = workbook.add_sheet('VIPOldman')
# 遍历字典数据并写入表中
for dic in myData:
    print(dic['name'])
    print(myData.index(dic))
    worksheet.write(myData.index(dic),0,dic['name'])
    worksheet.write(myData.index(dic),1,dic['nickName'])

fileName = "C:\\Users\\Administrator\\Desktop\\VIPOldman.xls"
# 保存文件
if(os.path.exists(fileName)==True):
    os.remove(fileName)
    print("删除了已存在的文件")
    workbook.save('C:\\Users\\Administrator\\Desktop\\VIPOldman.xls')
else:
    workbook.save('C:\\Users\\Administrator\\Desktop\\VIPOldman.xls')