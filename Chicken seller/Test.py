import numpy as np
import sys,os
import xlrd
import xlwt
from xlutils.copy import copy


def GetSheet(FileRb,FileWb,TargetName):
    
    Names = FileRb.sheet_names()
    for i in range(len(Names)):
        Sheet = FileWb.get_sheet(i)
        if(Names[i] == TargetName):
            return Sheet    
    return None

def MakePath(FileName,FileType):
    return os.path.dirname(sys.argv[0])+r'/ '[0] +str(FileName)+"."+FileType


def ReadExcel(Path):
    rb = xlrd.open_workbook(Path,formatting_info=True)
    return copy(rb),rb 

    
    
def FindColumnLength(FileRb,SheetName):

    Sheet = FileRb.sheet_by_name(SheetName)
    Column = Sheet.col_values(0)
    for i in range(len(Column)):
        if(Column[i]==""):
            return i+1
    
    return len(Column)    
    

#-----Main-----#

def DeleteSheet(FileWb,FileRb,SheetName):
    Names = FileRb.sheet_names()
    
    SheetFound = False
    for i in range(len(Names)):
        if(Names[i] == SheetName):
            SheetFound = True
            break
            
    if(SheetFound):
        del FileWb._Workbook__worksheets[i]
    else:
        print("can't delete:'"+SheetName+"', sheet not found!")
    


Path = MakePath("Data2","xls")
FileWb,FileRb = ReadExcel(Path)

DeleteSheet(FileWb,FileRb,"ItemSheet2")



FileWb.save(Path)






