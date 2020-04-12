from xlrd import open_workbook
from xlutils.copy import copy
import os
import re

def workstatistic(sour, dest):
    s = []
    num = 0
    files= os.listdir(sour)

    for file in files:
        s += re.findall(r"\d+\.?\d*",file)

    s = list(map(float, s))

    rexcel = open_workbook(dest)
    table0 = rexcel.sheets()[0]
    rows = rexcel.sheets()[0].nrows
    excel = copy(rexcel)
    table = excel.get_sheet(0)

    for i in range(rows):
        q = table0.cell(i, 0).value
        if q in s:
            table.write(i,2,1)
        else:
            table.write(i,2,0)

    excel.save(dest)

if __name__ == '__main__':
    workstatistic("E:/班级作业/操作系统选题/", 'E:/名单test/name.xls')