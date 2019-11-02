import os

# import numpy as np
import xlrd
import xlwt

# file = 'E:\\大学生新闻中心述职汇总'
# abspath = 'E:\\大学生新闻中心述职汇总\\'
# fileList = os.listdir(file)

file = '新媒体工作室2019第二学期述职报告评分表汇总'
abspath = '新媒体工作室2019第二学期述职报告评分表汇总\\'
fileList = os.listdir(file)

allscore = [[0 for col in range(2)] for row in range(16)]
filenum = 0


def read_excel(thisfile):
    print("正在读取 " + thisfile)
    wb = xlrd.open_workbook(filename=thisfile)  # 打开文件
    sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
    # print(sheet1.name, sheet1.nrows, sheet1.ncols)
    for i in range(sheet1.nrows):
        if i > 2:
            rows = sheet1.row_values(i)  # 获取行内容
            print(rows)
            # print(str(rows[0]) + ': ' + str(rows[7]))
            thisrow = i - 3
            allscore[thisrow][0] = str(rows[0])
            allscore[thisrow][1] = allscore[thisrow][1] + rows[7]


def write_excel():
    if 0 == filenum:
        print('路径下没有可分析文件!')
    else:
        print('正在生成评分表...')
        # 2. 创建Excel工作薄
        myWorkbook = xlwt.Workbook()
        # 3. 添加Excel工作表
        mySheet = myWorkbook.add_sheet('大学生新闻中心2019-11述职得分汇总')
        # 4. 写入数据
        for i in range(16):
            for j in range(2):
                mySheet.write(i, j, allscore[i][j])  # 写入数值
        # 5. 保存
        myWorkbook.save('评分表汇总排名\\大学生新闻中心2019-11述职得分汇总.xls')
        print('生成评分表成功！')


def delete_file():
    if os.path.exists('评分表汇总排名\\大学生新闻中心2019-11述职得分汇总.xls'):
        os.remove('评分表汇总排名\\大学生新闻中心2019-11述职得分汇总.xls')
        print('文件已删除')
    else:
        print("要删除的文件不存在！")


def calculate_score(num=None):
    for i in range(16):
        for j in range(2):
            if j == 1:
                allscore[i][j] = allscore[i][j] / num
    # print(allscore)


def sort(a):
    print('正在排序...')
    for k in range(len(a)):
        (a[k][0], a[k][1]) = (a[k][1], a[k][0])
    a.sort()
    for k in range(len(a)):
        (a[k][0], a[k][1]) = (a[k][1], a[k][0])
    print('排序成功!')


if __name__ == '__main__':
    delete_file()
    for m in range(len(fileList)):
        read_excel(abspath + fileList[m])
        filenum = m + 1
    print('文件读取成功！')
    calculate_score(filenum)
    sort(allscore)
    write_excel()

