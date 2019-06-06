import tools

#指标1-主函数
def main():
    a_file = '战略发展部经营分析201903.xlsx' #数据源表
    a_by_name_pl = '品类' #数据源sheet
    b_file = 'dm_jsc_fact_kpi_target_pl.xls' #数据存储表
    b_file_sheet = 'dm_jsc_fact_kpi_target_pl'

    #指标1-增量合并
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3459  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,8],[a+x,15],[a+x,92],[a+x,106],[a+x,127]] #进口烟--上海机场[2,8] 北京机场[2,15] 香港机场[2,92] 澳门机场[2,106] 游轮[2,127]
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            elif a_cell_pl_list[i] == -1e-06:
                a_cell_pl_list[i] == 0
            amount = a_cell_pl_list[i] + amount
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

    #指标1-合并
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3411   #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,1]] #进口烟--上海机场[2,8] 北京机场[2,15] 香港机场[2,92] 澳门机场[2,106] 游轮[2,127]
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            elif a_cell_pl_list[i] == -1e-06:
                a_cell_pl_list[i] == 0
            amount = a_cell_pl_list[i] + amount
        print(amount)
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
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
    b = 3411   #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,6]] #营收占比
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 9   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            elif a_cell_pl_list[i] == -1e-06:
                a_cell_pl_list[i] == 0
            amount = a_cell_pl_list[i] + amount
        print(amount)
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #ZZGXL
    a = 2   #获取数据的起始行数
    b = 3411   #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,7]] #营收增长贡献率
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 10   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            elif a_cell_pl_list[i] == -1e-06:
                a_cell_pl_list[i] == 0
            amount = a_cell_pl_list[i] + amount
        print(amount)
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

    #指标1-存量
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3411   #合并本年累计营收填报数据的起始行数
    c = 3459   #增量合并本年累计营收填报数据的起始行数
    d = 3420   #存量本年累计营收填报数据的起始行数
    for x in range(0,8):
        list1 = [[b+x,7]] #合并本年累计营收
        list2 = [[c+x,7]] #增量合并本年累计营收
        b_row = d+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list1 = tools.excel_cell_bylist(b_file,b_file_sheet,list1)   #获取合并本年累计营收
        a_cell_pl_list2 = tools.excel_cell_bylist(b_file,b_file_sheet,list2)   #获取增量合并本年累计营收
        amount = 0
        for i in range(len(a_cell_pl_list1)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            elif a_cell_pl_list[i] == -1e-06:
                a_cell_pl_list[i] == 0
            amount = a_cell_pl_list1[i] - a_cell_pl_list2[i]
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
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
    b = 3427   #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,15]] #首都机场本年累计营收
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            elif a_cell_pl_list[i] == -1e-06:
                a_cell_pl_list[i] == 0
            amount = a_cell_pl_list[i] + amount
        print(amount)
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

    #指标1-上海机场
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3435   #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,8]] #首都机场本年累计营收
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            elif a_cell_pl_list[i] == -1e-06:
                a_cell_pl_list[i] == 0
            amount = a_cell_pl_list[i] + amount
        print(amount)
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

    #指标1-上海机场
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3443   #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,92]] #首都机场本年累计营收
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            elif a_cell_pl_list[i] == -1e-06:
                a_cell_pl_list[i] == 0
            amount = a_cell_pl_list[i] + amount
        print(amount)
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

     #指标1-中免澳门
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3451   #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,106]] #首都机场本年累计营收
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            elif a_cell_pl_list[i] == -1e-06:
                a_cell_pl_list[i] == 0
            amount = a_cell_pl_list[i] + amount
        print(amount)
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

     #指标1-游轮
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3467   #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,127]] #首都机场本年累计营收
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            elif a_cell_pl_list[i] == -1e-06:
                a_cell_pl_list[i] == 0
            elif a_cell_pl_list[i] == 1e-06:
                a_cell_pl_list[i] = 0
            amount = a_cell_pl_list[i] + amount
        print(amount)
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
