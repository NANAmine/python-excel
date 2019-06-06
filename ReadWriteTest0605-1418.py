# -*- coding: utf-8 -*- 
# import  xdrlib ,sys
# import xlrd
# import xlwt
# from xlutils.copy import copy

# #打开excel文件
# def open_excel(file):
#     try:
#         data = xlrd.open_workbook(file)
#         return data
#     except Exception as e:
#         print(str(e))

# #获取单元格值
# def excel_cell_bycells(file,by_name,row,column):
#     data = open_excel(file) #打开excel文件
#     table = data.sheet_by_name(by_name) #根据sheet名字来获取excel中的sheet
#     cell = table.cell(row,column).value #获取cell值
#     return cell

# #获取多单元格值
# def excel_cell_bylist(file,by_name,list):
#     app = []
#     data = open_excel(file) #打开excel文件
#     table = data.sheet_by_name(by_name) #根据sheet名字来获取excel中的sheet
#     for x in range(len(list)):
#         row = list[x][0]
#         column = list[x][1]
#         cell = table.cell(row,column).value #获取cell值
#         app.append(cell)
#     return app

# #获取行
# def excel_table_byname(file, by_name, colnameindex):
#     data = open_excel(file) #打开excel文件
#     table = data.sheet_by_name(by_name) #根据sheet名字来获取excel中的sheet
#     colnames = table.row_values(colnameindex) #某一行数据 
#     app =[] #装读取结果的序列
#     if colnames: #如果行存在
#         for i in range(len(colnames)): #读取行的内容
#             app.append(row[i])
#     #获取指定行内容
#     print(app)
#     return app

# #根据名称获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的索引  ，by_name：Sheet1名称
# def excel_table_byname(file, by_name):
#     data = open_excel(file) #打开excel文件
#     table = data.sheet_by_name(by_name) #根据sheet名字来获取excel中的sheet
#     nrows = table.nrows #行数 
#     list =[] #装读取结果的序列
#     for rownum in range(0, nrows): #遍历每一行的内容
#          row = table.row_values(rownum) #根据行号获取行
#          if row: #如果行存在
#              app = [] #一行的内容
#              for i in range(len(row)): #一列列地读取行的内容
#                 app.append(row[i])
#              list.append(app) #装载数据
#     return list


# #将list中的内容写入一个新的file文件
# def testXlwt_list(file, Sheet, list):
    
#     book = xlwt.Workbook() #创建一个Excel
#     sheet1 = book.add_sheet(Sheet) #在其中创建一个名为hello的sheet
#     i = 0 #行序号
#     for app in list : #遍历list每一行
#         j = 0 #列序号
#         for x in app : #遍历该行中的每个内容（也就是每一列的）
#             sheet1.write(i, j, x) #在新sheet中的第i行第j列写入读取到的x值
#             j = j+1 #列号递增
#         i = i+1 #行号递增
#     book.save(file) #创建保存文件

# #指定位置写数据
# def testXlwt_cell(file, Sheet, row, column, cell):
#     book = xlrd.open_workbook(file) #打开一个Excel
#     xbook = copy(book)
#     sheet1 = xbook.get_sheet(Sheet) #在其中打开sheet
#     sheet1.write(row,column,cell) #往sheet里第row行第column列写一个数据
#     xbook.save(file) #保存文件

# #新建sheet页指定位置写数据
# def testrw(file):
#     book = xlwt.Workbook()
#     table = book.add_sheet('hello')
#     table.write(0,0,'5')
#     book.save(file)

import tools

#指标1-主函数
def main1():
    a_file = '副本战略发展部经营分析201904.xlsx' #数据源表
    a_by_name_pl = '品类' #数据源sheet
    b_file = 'dm_jsc_fact_kpi_target_pl.xls' #数据存储表
    b_file_sheet = 'dm_jsc_fact_kpi_target_pl'

    #指标1-增量合并
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3374  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,8],[a+x,15],[a+x,92],[a+x,106],[a+x,127]] #进口烟--上海机场[2,8] 北京机场[2,15] 香港机场[2,92] 澳门机场[2,106] 游轮[2,127]
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

    #指标1-合并
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3471   #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,1]] #进口烟--上海机场[2,8] 北京机场[2,15] 香港机场[2,92] 澳门机场[2,106] 游轮[2,127]
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        print(amount)
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #TBZF                                                                                                 香化和香食的问题，先保留置后
    # a = 2   #获取数据的起始行数
    # b = 3471   #填报数据的起始行数
    # for x in range(0,8):
    #     list = [[a+x,1]] #同比增幅
    #     b_row = b+x   #数据坐标行数
    #     print(b_row)
    #     b_column = 8   #数据坐标列数
    #     a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
    #     amount = 0    #坐标值求和
    #     for i in range(len(a_cell_pl_list)):
    #         amount = a_cell_pl_list[i] + amount
    #     print(amount)
    #     testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #ZB
    a = 2   #获取数据的起始行数
    b = 3471   #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,6]] #营收占比
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 9   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        print(amount)
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #ZZGXL
    a = 2   #获取数据的起始行数
    b = 3471   #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,7]] #营收增长贡献率
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 10   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        print(amount)
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

    #指标1-存量
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3471   #合并本年累计营收填报数据的起始行数
    c = 3374   #增量合并本年累计营收填报数据的起始行数
    d = 3382   #存量本年累计营收填报数据的起始行数
    for x in range(0,8):
        list1 = [[b+x,7]] #合并本年累计营收
        list2 = [[c+x,7]] #增量合并本年累计营收
        b_row = d+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list1 = excel_cell_bylist(b_file,b_file_sheet,list1)   #获取合并本年累计营收
        a_cell_pl_list2 = excel_cell_bylist(b_file,b_file_sheet,list2)   #获取增量合并本年累计营收
        amount = 0
        for i in range(len(a_cell_pl_list1)):
            amount = a_cell_pl_list1[i] - a_cell_pl_list2[i]
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #TBZF 由于香化和香食不同，                                                                                        先保留置后
    # a = 2   #获取数据的起始行数
    # b = 3471   #合并本年累计营收填报数据的起始行数
    # c = 3374   #增量合并本年累计营收填报数据的起始行数
    # d = 3382   #存量本年累计营收填报数据的起始行数
    # for x in range(0,8):
    #     list1 = [[b+x,7]] #合并本年累计营收
    #     list2 = [[c+x,7]] #增量合并本年累计营收
    #     b_row = d+x   #数据坐标行数
    #     print(b_row)
    #     b_column = 7   #数据坐标列数
    #     a_cell_pl_list1 = excel_cell_bylist(b_file,b_file_sheet,list1)   #获取合并本年累计营收
    #     a_cell_pl_list2 = excel_cell_bylist(b_file,b_file_sheet,list2)   #获取增量合并本年累计营收
    #     amount = 0
    #     for i in range(len(a_cell_pl_list1)):
    #         amount = a_cell_pl_list1[i] - a_cell_pl_list2[i]
    #     testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

    #指标1-北京机场
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3432   #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,15]] #首都机场本年累计营收
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        print(amount)
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

#指标2-主函数
def main2():
    a_file = '副本战略发展部经营分析201904.xlsx' #数据源表
    a_by_name_pl = '品类' #数据源sheet
    a_by_name_sntq = '上年同期'
    b_file = 'dm_jsc_fact_kpi_target_pl.xls' #数据存储表
    b_file_sheet = 'dm_jsc_fact_kpi_target_pl'

    #指标2-本部对外批发
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3448  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,148]] #本部对外批发
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #TBZF
    a = 2   #获取数据的起始行数
    b = 3448  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,151]] #本部对外批发
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 8   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            amount = a_cell_pl_list[i] + amount
            print(amount)
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

    #指标2-传统重点门店                                                                                              计算公式置后
    #


    #指标2-柬中免
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3406  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,99]] #本部对外批发
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #
    #TBZF
    a = 2   #获取数据的起始行数
    b = 3406  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,102]] #本部对外批发
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 8   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #SNTQZ                                                                                                       注意核对公式
    a = 164   #获取数据的起始行数
    b = 3406  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,148]] #上年同期值
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 22   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_sntq,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
            print(amount)
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #TBZF2                                                                                           取指标1存量同比增幅，香化食品置后
    

    #BNLJ2
    a = 2   #获取数据的起始行数
    b = 3471   #合并本年累计营收填报数据的起始行数
    c = 3374   #增量合并本年累计营收填报数据的起始行数
    d = 3406   #存量本年累计营收填报数据的起始行数
    for x in range(0,8):
        list1 = [[b+x,7]] #合并本年累计营收
        list2 = [[c+x,7]] #增量合并本年累计营收
        b_row = d+x   #数据坐标行数
        print(b_row)
        b_column = 24   #数据坐标列数
        a_cell_pl_list1 = excel_cell_bylist(b_file,b_file_sheet,list1)   #获取合并本年累计营收
        a_cell_pl_list2 = excel_cell_bylist(b_file,b_file_sheet,list2)   #获取增量合并本年累计营收
        amount = 0
        for i in range(len(a_cell_pl_list1)):
            amount = a_cell_pl_list1[i] - a_cell_pl_list2[i]
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

    #指标2-三亚店
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3510  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,36]] #本部对外批发
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)
    #TBZF
    a = 2   #获取数据的起始行数
    b = 3510  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,39]] #本部对外批发
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 8   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            amount = a_cell_pl_list[i] + amount
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #SNTQZ                                                                                                       注意核对公式
    a = 164   #获取数据的起始行数
    b = 3510  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,139]] #上年同期值
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 22   #数据坐标列数
        a_cell_pl_list = excel_cell_bylist(a_file,a_by_name_sntq,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
            print(amount)
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #TBZF2                                                                                           取指标1存量同比增幅，香化食品置后
    

    #BNLJ2
    a = 2   #获取数据的起始行数
    b = 3471   #合并本年累计营收填报数据的起始行数
    c = 3374   #增量合并本年累计营收填报数据的起始行数
    d = 3510   #存量本年累计营收填报数据的起始行数
    for x in range(0,8):
        list1 = [[b+x,7]] #合并本年累计营收
        list2 = [[c+x,7]] #增量合并本年累计营收
        b_row = d+x   #数据坐标行数
        print(b_row)
        b_column = 24   #数据坐标列数
        a_cell_pl_list1 = excel_cell_bylist(b_file,b_file_sheet,list1)   #获取合并本年累计营收
        a_cell_pl_list2 = excel_cell_bylist(b_file,b_file_sheet,list2)   #获取增量合并本年累计营收
        amount = 0
        for i in range(len(a_cell_pl_list1)):
            amount = a_cell_pl_list1[i] - a_cell_pl_list2[i]
        testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

if __name__=="__main__":
    #main1()
    main2()