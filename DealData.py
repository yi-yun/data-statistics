#-*- coding:utf-8 -*-
import math
import numpy as np
import matplotlib.pyplot as plt
import xlrd
import xlwt

info = dict()

# 计算 K 值
def cal(X,Y):
    for i in range(len(X)):
        X[i]=math.log(X[i])
    for i in range(len(Y)):
        Y[i]=math.log(Y[i])
    z1 = np.polyfit(X, Y, 1)
    return z1[0]

def read():
    data = xlrd.open_workbook('All.xlsx')
    table = data.sheets()[0]
    nrows = table.nrows  #行数
    # ncols = table.ncols
    res=[]
    for i in range(1,nrows): 
        line = table.row_values(i)
        list_X = []
        list_Y=[]
        for j in line[2::2]:
            list_X.append(j)
        for j in line[3::2]:
            list_Y.append(j)
        print(i)
        res.append(cal(list_X, list_Y))
    return res

def write():
    filename = xlwt.Workbook()
    sheet = filename.add_sheet("test")
    final()
    i=0
    for key, value in info.items():
        sheet.write(i, 0, key)
        sheet.write(i, 1,value)
        i+=1
    filename.save("test1.xls")
    
def read_test():
    data1 = xlrd.open_workbook('Employee2013.xlsx')
    data2 = xlrd.open_workbook('Employee2014.xlsx')
    data3 = xlrd.open_workbook('Employee2015.xlsx')
    data4 = xlrd.open_workbook('Employee2016.xlsx')
    data5 = xlrd.open_workbook('Employee2017.xlsx')
    data6 = xlrd.open_workbook('Employee201806.xlsx')

    table1 = data1.sheets()[0]
    table2 = data2.sheets()[0]
    table3 = data3.sheets()[0]
    table4 = data4.sheets()[0]
    table5 = data5.sheets()[0]
    table6 = data6.sheets()[0]
    
    col1 = table1.col_values(1)
    col2 = table2.col_values(1)
    col3 = table3.col_values(1)
    col4 = table4.col_values(1)
    col5 = table5.col_values(1)
    col6 = table6.col_values(1)
    for i in col5[2:]:
        sum = 1
        if i in col1[2:]:
            sum += 1
            find(i,table1)
        if i in col2[2:]:
            sum += 1
            find(i,table2)
        if i in col3[2:]:
            sum += 1
            find(i,table3)
        if i in col4[2:]:
            sum += 1
            find(i,table4)
        find(i,table5)
        if i in col6[2:]:
            sum += 1
            find(i,table6)

def find(i,table):
    for j in range(table.nrows):
        list=[]
        if i == table.row_values(j)[1]:
            list.append(table.row_values(j)[3])
            list.append(table.row_values(j)[4])
            if i in info:
                info[i].extend(list)
            else:
                info[i] = list

def final():
    read_test()
    for key, value in info.items():
        list_X = []
        list_Y=[]
        for j in value[0::2]:
            list_X.append(j)
        for j in value[1::2]:
            list_Y.append(j)
        info[key]=cal(list_X, list_Y)


# 以下为测试函数可有可无
# -------------------------
def draw():
    X = [28369, 29860, 32299, 32993, 31168, 32744]
    Y = [5218900.00, 7340700.0, 9616300.00, 10771500.00, 10578600.00, 8666400.00]
    for i in X:
        i = math.log(i)
    for i in Y:
        i = math.log(i)
    for i in range(len(X)):
        plt.plot([X[i]], [Y[i]], 'ro')
    # for i in range(len(X)):
    #     plt.plot([X[i]], [nihezhixian(X[i])],'go')
    plt.show()

def nihezhixian(x):
    return 980.56732 * x - 21933045.7

def a():
    data = xlrd.open_workbook('Test.xlsx')
    table = data.sheets()[0]
    nrows = table.nrows
    info = dict()
    for i in range(nrows):
        list=[]
        if "平安银行" == table.row_values(i)[1]:
            list.append(table.row_values(i)[3])
            list.append(table.row_values(i)[4])
            # print(list)
            if table.row_values(i)[1] in info:
                info[table.row_values(i)[1]].extend(list)
            else:
                info[table.row_values(i)[1]] = list
    print (info)
# ---------------------------------


if __name__ == '__main__':
    write()