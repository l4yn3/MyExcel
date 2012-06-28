#-*-coding:utf-8-*-
import os
import os.path
import xlrd
from openpyxl.reader.excel import load_workbook

class AbstractExcel:

    ename = None
    exists = True
    wb = None
    sh = None
    
    def __init__(self, excel_name):
        self.ename = excel_name
        self.CheckExists()
    
    def CheckExists(self):
        if not self.ename:
            self.exists = False
        if not os.path.isfile(self.ename):
            self.exists = False
    def OpenExcel(self):
        pass
    
    def GetSheet(self, num):
        pass
        
    def GetRowsNum(self):
        pass
        
    def GetColsNum(self):
        pass
    
    def GetCell(self, row, column):
        pass
    
class XlsExcel(AbstractExcel):
    def __init__(self, excel_name):
        AbstractExcel.__init__(self, excel_name)
        
    def OpenExcel(self):
        if not self.exists:
            return False;
        self.wb = xlrd.open_workbook(self.ename)
        return self
    
    def GetSheet(self, num):
        self.sh = self.wb.sheet_by_index(int(num))
        return self
    
    def GetRowsNum(self):
        return self.sh.nrows
    
    def GetColsNum(self):
        return self.sh.ncols
    
    def GetCell(self, row, column):
        return str(self.sh.cell(row, column).value)
            
class XlsxExcel(AbstractExcel):
    def __init__(self, excel_name):
        AbstractExcel.__init__(self, excel_name)
    
    def OpenExcel(self):
        if not self.exists:
            return False
        self.wb = load_workbook(self.ename)
        return self
    
    def GetSheet(self, num):
        self.sh = self.wb.get_sheet_by_name(self.wb.get_sheet_names()[int(num)])
        return self
    
    def GetRowsNum(self):
        return self.sh.get_highest_row()
    
    def GetColsNum(self):
        return self.sh.get_highest_col()
    
    def GetCell(self, _row, _column):
        return str(self.sh.cell(row = _row, column = _column).value)
        
class Factory:
    ename = None
    tools = {'xls':'XlsExcel', 'xlsx':'XlsxExcel'}
    def __init__(self, excel_name):
        self.ename = excel_name
    
    def GetExt(self):
        point_index = self.ename.rfind(".")
        return self.ename[point_index+1::]
    
    def Get(self):
        ext = self.GetExt()
        if ext == 'xls':
            return XlsExcel(self.ename)
        elif ext == 'xlsx':
            return XlsxExcel(self.ename)
        