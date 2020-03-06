import os
import xlrd
import xlwt

file = '大学生新闻中心值班表'
abspath = '大学生新闻中心值班表/'
fileList = os.listdir(file)
allname = [[0 for col in range(5)] for row in range(4)]


def read_excel_allinfo(thisfile):
    print("正在读取 " + thisfile + " 的空课表...")
    wb = xlrd.open_workbook(filename=thisfile)  # 打开文件
    sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
    print(sheet1.name, sheet1.nrows, sheet1.ncols)
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




if __name__ == '__main__':

    for filenum in range(len(fileList)):
        read_excel_allinfo(abspath + fileList[filenum])
        print(fileList[filenum])
    # write_excel()
