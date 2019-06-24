import numpy as np
import sys,os
import Classes as Cl
import xlrd
import xlwt
from xlutils.copy import copy
import Classes as Cl
import SizeAdjuster as Sa
import Classes as Cl


#Finds properties of bought items and puts it in excel sheet

#-----GlobalVariables-----#
StyleRed = xlwt.easyxf('pattern: pattern solid, fore_colour red')
StyleGreen = xlwt.easyxf('pattern: pattern solid, fore_colour green')
StyleBlue = xlwt.easyxf('pattern: pattern solid, fore_colour periwinkle')
#-----GlobalVariables-----#

#-----GlobalFunctions-----#
def ReadExcel(Path):
    rb = xlrd.open_workbook(Path,formatting_info=True)
    return copy(rb),rb 

def MakePath(FileName,FileType):
    return os.path.dirname(sys.argv[0])+r'/ '[0] +str(FileName)+"."+FileType
    
def GetSheet(FileRb,FileWb,TargetName):
    
    Names = FileRb.sheet_names()
    for i in range(len(Names)):
        Sheet = FileWb.get_sheet(i)
        if(Names[i] == TargetName):
            return Sheet
    
#-----GlobalFunctions-----#


def FindIndexOfId(TargetId,IdList):
    
    for i in range(len(IdList)):
        if(int(TargetId)==int(IdList[i])):
            return i
            
    return None


def MakePurchaseSheet(FileName="Data",PurchaseSheetName="PurchaseSheet",IdSheetName="IdSheet"):
    Path = MakePath(FileName,"xls")
    FileWb,FileRb = ReadExcel(Path)
    
    FileWb.set_colour_RGB(0x0A, 255, 80, 80)# Red
    FileWb.set_colour_RGB(0x11, 86, 160, 61)# Green
    
    IdSheet = FileRb.sheet_by_name(IdSheetName)
    PurchaseSheetRb = FileRb.sheet_by_name(PurchaseSheetName)
    PurchaseSheetWb = GetSheet(FileRb,FileWb,PurchaseSheetName)
    
    IdList = IdSheet.col_values(0)[1::]
    UrlList = IdSheet.col_values(2)[1::]
    
    PurchaseIdList = PurchaseSheetRb.col_values(0)[1::]
    NameList = PurchaseSheetRb.col_values(1)[1::]
    AmountList = PurchaseSheetRb.col_values(2)[1::]
    BoughtPriceList = PurchaseSheetRb.col_values(3)[1::]
    
    for i in range(len(PurchaseIdList)):
        Index = FindIndexOfId(PurchaseIdList[i],IdList)
        if(Index==None):
            print("Item Id not found. Skipping item...")
            continue
        Item = Cl.Item(NameList[i],UrlList[Index],1)
        
        PurchaseSheetWb.write(i+1,4,str(Item.CurrentPrice),)
        PurchaseSheetWb.write(i+1,7,str(Item.D3Rise))
        PurchaseSheetWb.write(i+1,8,str(Item.W1Rise))
        PurchaseSheetWb.write(i+1,9,str(Item.M1Rise))
        
        Investment = BoughtPriceList[i]*AmountList[i]
        Profit = (Item.CurrentPrice - BoughtPriceList[i])*AmountList[i]
        PurchaseSheetWb.write(i+1,6,str(int(Investment)))
        PurchaseSheetWb.write(i+1,5,str(int(Profit)))
        
    FileWb.save(Path)
        
        
MakePurchaseSheet()
    
            
        
    
    
    
    