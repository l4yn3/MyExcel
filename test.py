#-*-coding:utf-8-*-
import MyExcel
import sys

reload(sys)
sys.setdefaultencoding("utf-8")

sh = MyExcel.Factory("6.xlsx").Get().OpenExcel().GetSheet("0")
rows = sh.GetRowsNum()
for i in range(0, rows - 1):
    print sh.GetCell(i, 0) + "\r\n"
