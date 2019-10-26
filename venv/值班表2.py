import os
from typing import Counter

import xlrd
allname = [[0 for col in range(5)] for row in range(4)]
file = 'E:\\大学生新闻中心值班表'
fileList = os.listdir(file)
str1 = {'崔智鑫': 14, '崔若楠': 14, '刘岩': 9, '周振炜': 9, '辛千一': 8}
str3 = '刘岩 崔智鑫(13-16) 崔若楠(13-16) 方圆 陈丹'

# str = '向波 向波'
# print(str.split(' '))
s = str3.split(' ')
if '刘岩' in s:
    s.remove('刘岩')
    s1 = ''
    for i in s:
        s1 = s1 + ' ' + i
    print(s1)
# print(str1.keys())
file = 'E:\\大学生新闻中心值班表\\2019-2020第一学期后八周值班表1.xls'


def read_excel(thisfile):
    print("正在读取 " + thisfile + " 的空课表...")
    wb = xlrd.open_workbook(filename=thisfile)  # 打开文件
    sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
    print(sheet1.name, sheet1.nrows, sheet1.ncols)
    for i in range(4):
        name = []
        if i > 1:
            rows = sheet1.row_values(i)  # 获取行内容
            for j in range(5):
                if j > 0:
                    # print('i='+str(i)+'   j='+str(j))
                    if allname[i][j] != 0:
                        allname[i][j ] = (str(allname[i ][j ]).strip() + ' ' + rows[j].strip()).strip()
                    else:
                        allname[i ][j] = rows[j].strip()
            # print(rows)

    # print(sheet1.cell(3, 0).value)  # 获取表格里的内容，三种方式


def count_num():
    totalnum = [0]
    # 统计每个人空课数
    for i in range(len(allname)):
        for j in range(len(allname[i])):
            thisname = allname[i][j].__str__().split(' ')
            for m in range(len(thisname)):
                if totalnum[0] != 0:
                    totalnum[0] = (totalnum[0] + ' ' + thisname[m])
                else:
                    totalnum[0] = thisname[m]
    tempallname = totalnum[0].split(' ')
    print('----------元素出现次数----------')

    elem_time = eval((((str(Counter(tempallname)).split('Counter(')[1]) + "#").split(')#')[0]))
    print(elem_time)
    return elem_time


if __name__ == '__main__':
    read_excel(file)
    count_num()
