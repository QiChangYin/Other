# -*- coding: utf-8 -*-
import xdrlib, sys
import xlrd
import xlsxwriter
import shutil
import os
fileName = "E:\\qingsi\\销售明细报表_20180101至20180601(1-500).xlsx"
# writeTable = "E:\\qingsi\\result.xlsx"


insertRow = 1
insertCol = 0
tablesValue = []

hard_disk = "E:\\qingsi\\"
middleFilePath= hard_disk + "middleData\\"
rawFilePath= hard_disk + "rawData\\"
resultFilePath= hard_disk + "resultData\\"

def open_excel(file=fileName):
    data = xlrd.open_workbook(file)
    return data

# 根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def all_excel_table_byindex(file=fileName, colnameindex=0, by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数

    colnames = table.row_values(colnameindex)  # 某一行数据
    list = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            list.append(app)
    return list

# 根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file=middleFilePath, colnameindex=0, by_index=0):
    data = open_excel(middleFilePath+"mergeData.xls")
    table = data.sheets()[by_index]
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数

    colnames = table.row_values(colnameindex)  # 某一行数据
    list = []
    dataList = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        # print(row[1],row[3],row[4],int(row[8]))
        if row:
            app = [row[1],row[3],row[4],int(row[8])]
            gpp = [row[1],row[3],row[4]]
            # for i in range(len(colnames)):
                # app[colnames[i]] = row[i]
            list.append(app)
            dataList.append(gpp)
    return list,dataList


# 根据名称获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_name：Sheet1名称
def excel_table_byname(file=fileName, colnameindex=0, by_name=u'Sheet1'):
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows  # 行数
    colnames = table.row_values(colnameindex)  # 某一行数据
    list = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            list.append(app)
    return list

def merge_raw_table(biaotou):

    folder = os.path.exists(middleFilePath)
    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(middleFilePath)  # makedirs 创建文件时如果路径不存在会创建这个路径
        print ("--- Create new folder...  ---")
        print ("---  OK  ---")
    else:
        shutil.rmtree(middleFilePath)
        print ("---  Delete old folder And create new folder!  ---")

    # 在哪里搜索多个表格
    filelocation = hard_disk
    # 当前文件夹下搜索的文件名后缀
    fileform = "xlsx"
    # 将合并后的表格存放到的位置
    filedestination = middleFilePath
    # 合并后的表格命名为file
    file = "mergeData"

    # 首先查找默认文件夹下有多少文档需要整合
    import glob

    filearray = []
    for filename in glob.glob(filelocation + "*." + fileform):
        filearray.append(filename)
        # 以上是从pythonscripts文件夹下读取所有excel表格，并将所有的名字存储到列表filearray
    print("在默认文件夹下有%d个文档哦" % len(filearray))
    ge = len(filearray)
    matrix = [None] * ge
    # 实现读写数据

    # 下面是将所有文件读数据到三维列表cell[][][]中（不包含表头）
    import xlrd
    for i in range(ge):
        fname = filearray[i]
        bk = xlrd.open_workbook(fname)
        try:
            sh = bk.sheet_by_name("Sheet1")
        except:
            print("在文件%s中没有找到sheet1，读取文件数据失败,要不你换换表格的名字？" % fname)
        nrows = sh.nrows
        matrix[i] = [0] * (nrows - 1)

        ncols = sh.ncols
        for m in range(nrows - 1):
            matrix[i][m] = ["0"] * ncols

        for j in range(1, nrows):
            for k in range(0, ncols):
                matrix[i][j - 1][k] = sh.cell(j, k).value
                # 下面是写数据到新的表格test.xls中哦
    import xlwt
    filename = xlwt.Workbook()
    sheet = filename.add_sheet("mergeData")
    # 下面是把表头写上
    for i in range(0, len(biaotou)):
        sheet.write(0, i, biaotou[i])
        # 求和前面的文件一共写了多少行
    zh = 1
    for i in range(ge):
        for j in range(len(matrix[i])):
            for k in range(len(matrix[i][j])):
                sheet.write(zh, k, matrix[i][j][k])
            zh = zh + 1
    print("我已经将%d个文件合并成1个文件，并命名为%s.xls.快打开看看正确不？" % (ge, file))
    filename.save(filedestination + file + ".xls")


def repeat_field(table_key_value,long_table_list,value):
    n = len(table_key_value)
    row_value = 0
    for i in range(n):
        if table_key_value[i] == value:
            row_value = long_table_list[i][3]+row_value
    print(value,row_value)
    return row_value

def main():

    # 获取表头数据
    tablesValueList = all_excel_table_byindex()
    keys = list(tablesValueList[0].keys())
    print(keys)

    #合并原始表数据到一个中间表
    merge_raw_table(keys)

    # folder_exist = os.path.exists(resultFilePath)
    # if folder_exist:
    #     shutil.rmtree(resultFilePath)

    workbook = xlsxwriter.Workbook(resultFilePath+'result.xlsx')
    worksheet = workbook.add_worksheet()
    # Add a bold format to use to highlight cells. 设置粗体，默认是False
    bold = workbook.add_format({'bold': True})
    # Write some data headers. 带自定义粗体blod格式写表头
    worksheet.write('A1', '货号', bold)
    worksheet.write('B1', '颜色', bold)
    worksheet.write('C1', 'M', bold)
    worksheet.write('D1', 'L', bold)
    worksheet.write('E1', 'XL', bold)
    worksheet.write('F1', 'XXL', bold)
    worksheet.write('G1', '3XL', bold)
    worksheet.write('H1', 'Other', bold)
    worksheet.write('I1', '总计', bold)

    LongTablesValue, keyTablesValue = excel_table_byindex()

    print(LongTablesValue)
    insertRow = 1
    insertCol = 0
    a = {}
    for i,row in enumerate(keyTablesValue):
        if keyTablesValue.count(row) >= 1:
            a[str(row)] = repeat_field(keyTablesValue,LongTablesValue,keyTablesValue[i])
            print(a[str(row)])

    print(a)


    b = dict()
    for key, value in a.items():
        ca= key.split(",")
        number = ca[0][2:-1]
        color = ca[1][2:-1]
        type =  ca[2][2:-2]
        tkey = str(number + ":" + color )
        if  tkey in b.keys():
            print("重复项")
        else:
            b[tkey] = {
                "M":0,
                "L":0,
                "XL":0,
                "XXL":0,
                "3XL": 0,
                "other":0
            }
        print(tkey,value)

        if type == "M":
            j = 2
            b[tkey]["M"] = int(b[tkey]["M"]) + int(value)
            # print(b[tkey]["M"] )
        elif type == "L":
            j = 3
            b[tkey]["L"] = int(value) + int(b[tkey]["L"])
        elif type == "XL":
            j = 4
            b[tkey]["XL"] = int(value) + int(b[tkey]["XL"])
        elif type == "XXL":
            b[tkey]["XXL"] = int(value) + int(b[tkey]["XXL"])
            j = 5
        elif type == "3XL":
            b[tkey]["3XL"] = int(value) + int(b[tkey]["3XL"])
            j = 6
        else:
            b[tkey]["other"] = int(value) + int(b[tkey]["other"])
            j = 7
    for key, value in b.items():
        # print(key,value["M"],value["L"],value["XL"],value["XXL"],value["other"])
        if int(value["M"]+value["L"]+value["XL"]+value["XXL"]+value["3XL"]+value["other"]) != 0:
            cao = key.split(':')
            worksheet.write(insertRow, insertCol, cao[0])  # 带默认格式写入
            worksheet.write(insertRow, insertCol+1, cao[1])  # 带默认格式写入
            worksheet.write(insertRow, insertCol + 2, value["M"] if int(value["M"]) > 0 else " ")  # 带默认格式写入
            worksheet.write(insertRow, insertCol + 3, value["L"] if int(value["L"]) > 0 else " ")  # 带默认格式写入
            worksheet.write(insertRow, insertCol + 4, value["XL"] if int(value["XL"]) > 0 else " ")  # 带自定义money格式写入
            worksheet.write(insertRow, insertCol + 5, value["XXL"] if int(value["XXL"]) > 0 else " ")  # 带自定义money格式写入
            worksheet.write(insertRow, insertCol + 6, value["3XL"] if int(value["3XL"]) > 0 else " ")  # 带自定义money格式写入
            worksheet.write(insertRow, insertCol + 7, value["other"] if int(value["other"]) > 0 else " ")  # 带自定义money格式写入
            worksheet.write(insertRow, insertCol + 8, value["M"] + value["L"] + value["XL"] + value["XXL"] + value["3XL"] +  value["other"])  # 带自定义money格式写入
            insertRow += 1
    workbook.close()


if __name__ == "__main__":
    main()

