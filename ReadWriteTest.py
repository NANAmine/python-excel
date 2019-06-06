# -*- coding: utf-8 -*- 
import  xdrlib ,sys
import xlrd
import xlwt


#打开excel文件
def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))

#获取单元格值
def excel_cell_bycells(file,by_name,row,column):
    data = open_excel(file) #打开excel文件
    table = data.sheet_by_name(by_name) #根据sheet名字来获取excel中的sheet
    cell = table.cell(row,column).value #获取cell值
    return cell

#获取行
def excel_table_byname(file, by_name, colnameindex):
    data = open_excel(file) #打开excel文件
    table = data.sheet_by_name(by_name) #根据sheet名字来获取excel中的sheet
    colnames = table.row_values(colnameindex) #某一行数据 
    app =[] #装读取结果的序列
    if colnames: #如果行存在
        for i in range(len(colnames)): #读取行的内容
            app.append(row[i])
    #获取指定行内容
    print(app)
    return app

#根据名称获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的索引  ，by_name：Sheet1名称
def excel_table_byname(file, by_name):
    data = open_excel(file) #打开excel文件
    table = data.sheet_by_name(by_name) #根据sheet名字来获取excel中的sheet
    nrows = table.nrows #行数 
    list =[] #装读取结果的序列
    for rownum in range(0, nrows): #遍历每一行的内容
         row = table.row_values(rownum) #根据行号获取行
         if row: #如果行存在
             app = [] #一行的内容
             for i in range(len(row)): #一列列地读取行的内容
                app.append(row[i])
             list.append(app) #装载数据
    return list

#根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的索引  ，by_index：表的索引
# def excel_table_byindex(file= 'test.xlsx',colnameindex=0,by_index=0):
#     data = open_excel(file) #打开excel文件
#     table = data.sheets()[by_index] #根据sheet序号来获取excel中的sheet
#     nrows = table.nrows #行数
#     ncols = table.ncols #列数
#     colnames =  table.row_values(colnameindex) #某一行数据 
#     list =[] #装读取结果的序列
#     for rownum in range(0,nrows): #遍历每一行的内容

#          row = table.row_values(rownum) #根据行号获取行
#          if row: #如果行存在
#              app = [] #一行的内容
#              for i in range(len(colnames)): #一列列地读取行的内容
#                 app.append(row[i])
#              if app[0] == app[1] : #如果这一行的第一个和第二个数据相同才将其装载到最终的list中
#              	list.append(app)
#     testXlwt('new.xlsx', list) #调用写函数，讲list内容写到一个新文件中
#     return list



#将list中的内容写入一个新的file文件
def testXlwt_list(file, Sheet, list):
	
	book = xlwt.Workbook() #创建一个Excel
	sheet1 = book.add_sheet(Sheet) #在其中创建一个名为hello的sheet
	i = 0 #行序号
	for app in list : #遍历list每一行
		j = 0 #列序号
		for x in app : #遍历该行中的每个内容（也就是每一列的）
			sheet1.write(i, j, x) #在新sheet中的第i行第j列写入读取到的x值
			j = j+1 #列号递增
		i = i+1 #行号递增
	# sheet1.write(0,0,'cloudox') #往sheet里第一行第一列写一个数据
	# sheet1.write(1,0,'ox') #往sheet里第二行第一列写一个数据
	book.save(file) #创建保存文件

#指定位置写数据
def testXlwt_cell(file, Sheet, row, column, cell):
    book = xlwt.Workbook() #创建一个Excel
    sheet1 = book.add_sheet(Sheet) #在其中创建一个名为hello的sheet
    sheet1.write(row,column,cell) #往sheet里第row行第column列写一个数据
    book.save(file) #创建保存文件

#新建sheet页指定位置写数据
def testrw(file):
	book = xlwt.Workbook()
	table = book.add_sheet('hello')
	table.write(0,0,'5')
	book.save(file)

#主函数
def main():
    file = '副本战略发展部经营分析201904.xlsx' #数据源表
    by_name = '上年同期' #数据源sheet
    file2 = 'test2.xls' #数据存储表
    testrw('new2.xls')
    #tables = excel_table_byindex()
    tables = excel_table_byname(file,by_name)
    #将数据写入到的Excel文件中
    testXlwt_list(file2, by_name, tables)
    for column in tables:
        print(column)

   # tables = excel_table_byname()
   # for row in tables:
   #     print row

   # testrw()

if __name__=="__main__":
    main()