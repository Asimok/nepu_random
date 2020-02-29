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

def read_excel_getNameAndMajor(thisfile):
    print("正在读取 " + thisfile)
    wb = xlrd.open_workbook(filename=thisfile)  # 打开文件
    sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
    for i in range(sheet1.nrows):
        if i > 0:
            rows = sheet1.row_values(i)  # 获取行内容
            nameAndMajor.append(rows[0]+'('+rows[3]+')')
            name.append(rows[0])



def read_excel_allinfo(thisfile):
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
    print("正在统计推送次数")
    global nameAndSorcePartInDict
    global nameAndSorcePushDict
    times=[]
    for i in range(name.__len__()):
        times.append(0)
    nameAndSorcePartInDict=dict(zip(nameAndMajor,times))
    nameAndSorcePushDict = dict(zip(nameAndMajor, times))
    for key in nameAndSorcePartInDict:
        nameAndSorcePartInDict[key]=nameAndSorcePartInDict[key]+PartIn.count(key[0:key.index('(')])
        nameAndSorcePushDict[key]=nameAndSorcePushDict[key]+Push.count(key[0:key.index('(')])



def sortDict():
    print("正在排序")
    global nameAndSorcePartInDict
    global nameAndSorcePushDict
    global sort_PartInDict
    global sort_PushDict
    sort_PartIn = sorted(nameAndSorcePartInDict.items(), key=lambda item: item[1], reverse=True)
    sort_Push = sorted(nameAndSorcePushDict.items(), key=lambda item: item[1], reverse=True)
    sort_PartInDict=dict(sort_PartIn)
    sort_PushDict=dict(sort_Push)
    # print("参与情况")
    # print(sort_PartInDict)
    # print("推送情况")
    # print(sort_PushDict)


def all():
    print("正在汇总")
    global sort_PartInDict
    global sort_PushDict
    global sort_end
    global sort_end_dict
    for key in sort_PartInDict:
        # print(key)
        sort_PushDict[key]=int(sort_PushDict[key])+int(sort_PartInDict[key])
    sort_end = sorted(sort_PushDict.items(), key=lambda item: item[1], reverse=True)
    sort_end_dict=dict(sort_end)
    print(sort_end_dict)

def write_excel():
        global sort_end
        print('正在生成汇总表...')
        # 2. 创建Excel工作薄
        myWorkbook = xlwt.Workbook()
        # 3. 添加Excel工作表
        mySheet = myWorkbook.add_sheet('大学生新闻中心2020量化统计汇总')
        # 4. 写入数据
        i=0
        for key in sort_end_dict:
            mySheet.write(i, 0, key)  # 写入数值
            mySheet.write(i, 1,  sort_end_dict[key])  # 写入数值
            i=i+1
        # 5. 保存
        myWorkbook.save('大学生新闻中心2020量化统计汇总.xls')
        print('生成汇总表成功！')


def delete_file():
    if os.path.exists('大学生新闻中心2020量化统计汇总.xls'):
        os.remove('大学生新闻中心2020量化统计汇总.xls')
        print('文件已删除')
    else:
        print("要删除的文件不存在！")


if __name__ == '__main__':

    read_excel_getNameAndMajor("专业统计.xlsx")
    read_excel_allinfo("大学生新闻中心量化数据统计表.xlsx")
    countNum()
    sortDict()
    all()
    delete_file()
    write_excel()

