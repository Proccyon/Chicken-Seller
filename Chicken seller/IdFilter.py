import numpy as np
import sys,os
import Classes as Cl
import xlrd
import xlwt
from xlutils.copy import copy
import Classes as Cl
import time
import SizeAdjuster as Sa




def ReadExcel(Path):
    rb = xlrd.open_workbook(Path,formatting_info=True)
    return copy(rb),rb 

def MakePath(FileName,FileType):
    return os.path.dirname(sys.argv[0])+r'/ '[0] +str(FileName)+"."+FileType


def FindColumn(Sheet,ColumnName):
    
    for i in range(len(Sheet.row_values(0))):
        Column = Sheet.col_values(i)
        if(ColumnName == Column[0]):
            return Column

def ConvertBuyLimit(BuyLimit):
    if(BuyLimit==""):
        return 1
    else:
        return int(float((str(BuyLimit).replace(",",""))))



def MakeFilteredIdSheet(FileName="Data",IdSheetName = "IdSheet",ItemSheetName="ItemSheet",FilteredSheetName="FilteredIdSheet2",PriceLimit=800000):
    Path = MakePath(FileName,"xls")
    FileWb,FileRb = ReadExcel(Path)
    
    ItemSheet = FileRb.sheet_by_name(ItemSheetName)
    IdSheet = FileRb.sheet_by_name(IdSheetName)
    
    IdList = IdSheet.col_values(0)
    NameList = IdSheet.col_values(1)
    UrlList = IdSheet.col_values(2)
    BuyLimitList = IdSheet.col_values(3)  
    PriceList = FindColumn(ItemSheet,"Price")

    FilteredSheet = Sa.FitSheetWrapper(FileWb.add_sheet(FilteredSheetName))
    
    index = 1
    
    FilteredSheet.write(0,0,"Id")
    FilteredSheet.write(0,1,"Item Name")
    FilteredSheet.write(0,2,"Url")
    FilteredSheet.write(0,3,"Buy Limit")
    
    for i in range(1,len(IdList)):
        
        BuyLimit = ConvertBuyLimit(BuyLimitList[i])    
        Price = int(PriceList[i])
        
        if(Price*BuyLimit >= PriceLimit):
            FilteredSheet.write(index,0,IdList[i])
            FilteredSheet.write(index,1,NameList[i])
            FilteredSheet.write(index,2,UrlList[i])
            FilteredSheet.write(index,3,str(BuyLimit))
            index += 1
        
    FileWb.save(Path)
            
#-----Main-----#

PriceLimit = 800000
MakeFilteredIdSheet(PriceLimit=PriceLimit)
    
    
#-----Main-----#
    
