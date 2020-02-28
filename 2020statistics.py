import os

import xlrd
import xlwt
PartIn=''
Push=''
nameAndMajor=[]
name=[]
# 参与次数
nameAndSorcePartInDict ={}
# 推送次数
nameAndSorcePushDict ={}

def getNameAndMajor(thisfile):
    print("正在读取 " + thisfile)
    wb = xlrd.open_workbook(filename=thisfile)  # 打开文件
    sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
    for i in range(sheet1.nrows):
        if i > 0:
            rows = sheet1.row_values(i)  # 获取行内容
            nameAndMajor.append(rows[0]+'('+rows[3]+')')
            name.append(rows[0])



def read_excel(thisfile):
    global PartIn
    global Push
    print("正在读取 " + thisfile)
    wb = xlrd.open_workbook(filename=thisfile)  # 打开文件
    sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
    for i in range(sheet1.nrows):
        if i > 1:
            rows = sheet1.row_values(i)  # 获取行内容
            # print(rows)
            PartIn = PartIn.replace('，', '')
            PartIn = PartIn.replace('、', '')
            PartIn = PartIn.replace('/', '')
            PartIn=PartIn+''.join(rows[3].split())+''.join(rows[4].split())+\
                   ''.join(rows[5].split())+''.join(rows[6].split())
            Push=Push+''.join(rows[9].split())
            # print(rows)

def countNum():
    times=[]
    for i in range(name.__len__()):
        times.append(0)
    nameAndSorcePartInDict=dict(zip(name,times))
    nameAndSorcePushDict = dict(zip(name, times))
    for key in nameAndSorcePartInDict:
        # print(key)
        nameAndSorcePartInDict[key]=nameAndSorcePartInDict[key]+PartIn.count(key)
        nameAndSorcePushDict[key]=nameAndSorcePushDict[key]+Push.count(key)
    print("参与情况")
    print(nameAndSorcePartInDict)
    print("推送情况")
    print(nameAndSorcePushDict)


def write_excel():
        print('正在生成评分表...')
        # 2. 创建Excel工作薄
        myWorkbook = xlwt.Workbook()
        # 3. 添加Excel工作表
        mySheet = myWorkbook.add_sheet('大学生新闻中心2019-11述职得分汇总')
        # 4. 写入数据
        # for i in range(len(allscore)):
        #     for j in range(2):
        #         mySheet.write(i, j, allscore[i][j])  # 写入数值
        # 5. 保存
        myWorkbook.save('评分表汇总排名\\大学生新闻中心2019-11述职得分汇总.xls')
        print('生成评分表成功！')


def delete_file():
    if os.path.exists('评分表汇总排名\\大学生新闻中心2019-11述职得分汇总.xls'):
        os.remove('评分表汇总排名\\大学生新闻中心2019-11述职得分汇总.xls')
        print('文件已删除')
    else:
        print("要删除的文件不存在！")


if __name__ == '__main__':

    getNameAndMajor("专业统计.xlsx")
    # delete_file()
    read_excel("大学生新闻中心量化数据统计表.xlsx")
    # print('文件读取成功！')
    #
    # write_excel()
    # print(name)
    # print(PartIn)
    # print(Push)
    countNum()

