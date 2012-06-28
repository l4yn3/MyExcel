#-*-coding:utf-8-*-
import MyExcel
import sys

reload(sys)
sys.setdefaultencoding("utf-8")

sh = MyExcel.Factory("hi.xlsx").Get().OpenExcel().GetSheet("0")
for i in range(0, sh.GetRowsNum()):
    print sh.GetCell(i, 0), "\r\n"
#sh = MyExcel.Factory("hi.xls").Get().CreateExcel().CreateSheet("0")
#for i in range(0, 10):
#    sh.SetCell(i, 0, str(i))
#sh.Save()

print 'ok'
