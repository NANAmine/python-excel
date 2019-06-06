import tools

#指标2-主函数
def main():
    a_file = '战略发展部经营分析201904.xlsx' #数据源表
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
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #TBZF
    a = 2   #获取数据的起始行数
    b = 3448  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,151]] #本部对外批发
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 8   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            amount = a_cell_pl_list[i] + amount
            print(amount)
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

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
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #
    #TBZF
    a = 2   #获取数据的起始行数
    b = 3406  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,102]] #本部对外批发
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 8   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #SNTQZ                                                                                                       注意核对公式
    a = 164   #获取数据的起始行数
    b = 3406  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,148]] #上年同期值
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 22   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_sntq,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
            print(amount)
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
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
        a_cell_pl_list1 = tools.excel_cell_bylist(b_file,b_file_sheet,list1)   #获取合并本年累计营收
        a_cell_pl_list2 = tools.excel_cell_bylist(b_file,b_file_sheet,list2)   #获取增量合并本年累计营收
        amount = 0
        for i in range(len(a_cell_pl_list1)):
            amount = a_cell_pl_list1[i] - a_cell_pl_list2[i]
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

    #指标2-三亚店
    #BNLJ
    a = 2   #获取数据的起始行数
    b = 3510  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,36]] #本部对外批发
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 7   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)
    #TBZF
    a = 2   #获取数据的起始行数
    b = 3510  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,39]] #本部对外批发
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 8   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_pl,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            if a_cell_pl_list[i] == "":
                a_cell_pl_list[i] = 0
            amount = a_cell_pl_list[i] + amount
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
    #SNTQZ                                                                                                       注意核对公式
    a = 164   #获取数据的起始行数
    b = 3510  #填报数据的起始行数
    for x in range(0,8):
        list = [[a+x,139]] #上年同期值
        b_row = b+x   #数据坐标行数
        print(b_row)
        b_column = 22   #数据坐标列数
        a_cell_pl_list = tools.excel_cell_bylist(a_file,a_by_name_sntq,list)   #获取多组坐标值
        amount = 0    #坐标值求和
        for i in range(len(a_cell_pl_list)):
            amount = a_cell_pl_list[i] + amount
            print(amount)
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表
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
        a_cell_pl_list1 = tools.excel_cell_bylist(b_file,b_file_sheet,list1)   #获取合并本年累计营收
        a_cell_pl_list2 = tools.excel_cell_bylist(b_file,b_file_sheet,list2)   #获取增量合并本年累计营收
        amount = 0
        for i in range(len(a_cell_pl_list1)):
            amount = a_cell_pl_list1[i] - a_cell_pl_list2[i]
        tools.testXlwt_cell(b_file, b_file_sheet, b_row, b_column, amount)   #数据写入报表

if __name__=="__main__":
    main2()