# 大学生新闻中心1-8周空课表
import os

import xlrd
import xlwt

file = '大学生新闻中心1-8周空课表'
abspath = '大学生新闻中心1-8周空课表\\'
fileList = os.listdir(file)
week = ["周一", "周二", "周三", "周四", "周五"]
curricula = ["第一大节", "第二大节", "第三大节", "第四大节", "第五大节", "第六大节"]
all_Curriculum = [['' for col in range(6)] for row in range(8)]


def read_excel_Curriculum(thisfile):
    print("正在读取 " + thisfile + " 的空课表...")
    wb = xlrd.open_workbook(filename=thisfile)  # 打开文件
    sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
    for i in range(8):
        if i > 1:
            rows = sheet1.row_values(i)  # 获取行内容
            for j in range(6):
                if j > 0:
                    if str(rows[j]).__contains__("周"):
                        continue
                    else:
                        all_Curriculum[i][j] = all_Curriculum[i][j].strip()+ " " + rows[j].strip()


def write_excel():
    print('正在汇总空课表...')
    # 2. 创建Excel工作薄
    myWorkbook = xlwt.Workbook()
    # 3. 添加Excel工作表
    mySheet = myWorkbook.add_sheet('大学生新闻中心2020前八周空课表')
    # 4. 写入数据
    for i in range(8):
        for j in range(6):
            mySheet.write(i, j, all_Curriculum[i][j])  # 写入数值
    # 5. 保存
    myWorkbook.save('导出\\2020第一学期前八周空课表.xls')
    print("已生成")

def generateForm():
    for i in range(8):
        if i>1:
            all_Curriculum[i][0]=curricula[i-2]
    for j in range(6):
        if j>0:
            all_Curriculum[1][j] = week[j - 1]

if __name__ == '__main__':
    # for filenum in range(len(fileList)):
    for filenum in range(len(fileList)):
        read_excel_Curriculum(abspath + fileList[filenum])
    generateForm()
    write_excel()

    # for i in range(8):
    #     print(all_Curriculum[i])
