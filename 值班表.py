import datetime
# coding=utf-8
import os
import xlrd
import xlwt
from collections import Counter

# file = 'E:\\大学生新闻中心值班表'
# abspath = 'E:\\大学生新闻中心值班表\\'
# fileList = os.listdir(file)
# file1 = 'E:\\大学生新闻中心值班表\\test.xlsx'
# file2 = 'E:\\大学生新闻中心值班表\\test2.xlsx'
file = '大学生新闻中心值班表'
abspath = '大学生新闻中心值班表\\'
fileList = os.listdir(file)

allname = [[0 for col in range(5)] for row in range(4)]


def read_excel_allinfo(thisfile):
    print("正在读取 " + thisfile + " 的空课表...")
    wb = xlrd.open_workbook(filename=thisfile)  # 打开文件
    sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
    # print(sheet1.name, sheet1.nrows, sheet1.ncols)
    for i in range(6):
        name = []
        if i > 1:
            rows = sheet1.row_values(i)  # 获取行内容
            for j in range(6):
                if j > 0:
                    # print('i='+str(i)+'   j='+str(j))
                    if allname[i - 2][j - 1] != 0:
                        allname[i - 2][j - 1] = (str(allname[i - 2][j - 1]).strip() + ' ' + rows[j].strip()).strip()
                    else:
                        allname[i - 2][j - 1] = rows[j].strip()
            # print(rows)

    # print(sheet1.cell(3, 0).value)  # 获取表格里的内容，三种方式


def write_excel():
    print('正在生成值班表...')
    # 2. 创建Excel工作薄
    myWorkbook = xlwt.Workbook()
    # 3. 添加Excel工作表
    mySheet = myWorkbook.add_sheet('大学生新闻中心值班表')
    # 4. 写入数据
    for i in range(4):
        for j in range(5):
            mySheet.write(i, j, allname[i][j])  # 写入数值
    # 5. 保存
    # myWorkbook.save('E:\\大学生新闻中心值班表\\2019-2020第一学期后八周值班表1.xls')
    myWorkbook.save('导出\\2019-2020第一学期后八周值班表.xls')

def get_table():
    print('----------临时值班表----------')
    for m in range(len(allname)):
        print(allname[m])


# def change_num():
# 优化 allnum 列表数量

def count_num():
    totalnum = [0]
    # 统计每个人空课数
    for i in range(len(allname)):
        for j in range(len(allname[i])):
            thisname = allname[i][j].split(' ')
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


def check_oneday():
    # 每人一天只出现一次 删除第二次及以后出现的
    # 统计每人一天内空课数
    totalnum = [0, 0, 0, 0, 0]
    everydayname = [0, 0, 0, 0, 0]
    for i in range(len(allname)):
        for j in range(len(allname[i])):
            thisname = allname[i][j].split(' ')
            for m in range(len(thisname)):
                if totalnum[j] != 0:
                    totalnum[j] = (totalnum[j] + ' ' + thisname[m])
                else:
                    totalnum[j] = thisname[m]
    tempallname = totalnum
    print('----------每人一天内空课数----------')
    for i in range(len(tempallname)):
        print(Counter(tempallname[i].__str__().split(' ')))
        elem_time = eval(
            (((str(Counter(tempallname[i].__str__().split(' '))).split('Counter(')[1]) + "#").split(')#')[0]))
        everydayname[i] = elem_time
        print('周' + str(i + 1) + '  ' + str(everydayname[i]))

    # 每人一天只出现一次 删除第二次及以后出现的
    for lengthofday in range(len(allname)):
        for key, value in everydayname[lengthofday].items():
            # print(str(key) + ':' + str(value))
            if value > 1:
                willdeletename = key
                i = 0
                for time in range(len(allname)):
                    for j in allname[time][lengthofday].__str__().split(' '):
                        if j.__eq__(willdeletename):
                            i = i + 1
                            if i > 1:
                                # 删除第二次开始的
                                # print(willdeletename + "  出现 第" + str(i) + ' 次  将要删除')
                                s = allname[time][lengthofday].split(' ')
                                s.remove(willdeletename)
                                s1 = ''
                                for namenum in s:
                                    s1 = s1 + ' ' + namenum
                                allname[time][lengthofday] = s1.strip()
                                # print(allname[time][lengthofday])


def delete_maxfour():
    # 每节课值班人数不超过4人 超过四人删除出现次数最多的
    totalnum = [0]
    # 统计每个人空课数
    for i in range(len(allname)):
        for j in range(len(allname[i])):
            thisname = allname[i][j].split(' ')
            for m in range(len(thisname)):
                if totalnum[0] != 0:
                    totalnum[0] = (totalnum[0] + ' ' + thisname[m])
                else:
                    totalnum[0] = thisname[m]
    tempallname = totalnum[0].split(' ')
    print('----------元素出现次数----------')

    elem_time = eval((((str(Counter(tempallname)).split('Counter(')[1]) + "#").split(')#')[0]))

    # 每节课值班人数不超过4人 超过四人删除出现次数最多的
    for i in range(len(allname)):
        for j in range(len(allname[i])):
            while len(allname[i][j].split(' ')) > 6:
                s2 = allname[i][j].split(' ')
                for containnumname in s2:
                    if str(containnumname).__contains__('1'):
                        if len(allname[i][j].split(' ')) > 6:
                            s3 = allname[i][j].split(' ')
                            s3.remove(containnumname)
                            s4 = ''
                            for namenum in s3:
                                s4 = s4 + ' ' + namenum
                            allname[i][j] = s4.strip()
                            print(containnumname + '  低效利用删除')
                for willdelete in elem_time.keys():
                    # print('超过4人')
                    if willdelete in allname[i][j].split(' '):
                        if len(allname[i][j].split(' ')) > 4:
                            print(willdelete)
                            s = allname[i][j].split(' ')
                            s.remove(willdelete)
                            s1 = ''
                            for namenum in s:
                                s1 = s1 + ' ' + namenum
                            allname[i][j] = s1.strip()
                            print(willdelete + '  高频删除')
                        else:
                            continue






def delete_maxtwo():
    # 每人每周最多值班2次 超过两次删除人多的时候

    totalnum = [0]
    # 统计每个人空课数
    for i in range(len(allname)):
        for j in range(len(allname[i])):
            thisname = allname[i][j].split(' ')
            for m in range(len(thisname)):
                if totalnum[0] != 0:
                    totalnum[0] = (totalnum[0] + ' ' + thisname[m])
                else:
                    totalnum[0] = thisname[m]
    tempallname = totalnum[0].split(' ')
    print('----------元素出现次数----------')

    elem_time = eval((((str(Counter(tempallname)).split('Counter(')[1]) + "#").split(')#')[0]))
    print(elem_time)



if __name__ == '__main__':
    # 读取Excel
    # read_excel_allinfo(file1)
    # read_excel_allinfo(file2)
    for filenum in range(len(fileList)):
        read_excel_allinfo(abspath + fileList[filenum])

    count_num()
    check_oneday()
    print(allname[0][0])
    delete_maxfour()
    # delete_maxtwo()
    count_num()
    get_table()
    write_excel()
