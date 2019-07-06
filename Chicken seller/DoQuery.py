import numpy as np
import sys,os
import xlrd
import xlwt
from xlutils.copy import copy

import Classes as Cl
import SizeAdjuster as Sa


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

#Creates an item using information of a single row of "PriceSheet"
def CreateItem(RowNumber,ParameterList,FileRb,SheetName):

    PriceSheet = FileRb.sheet_by_name(SheetName)
    Row = PriceSheet.row_values(RowNumber)
    ItemName = Row[0]
    BuyLimit = Row[1]
    PriceList = Row[2:]
    PriceList = [int(i) for i in PriceList]
    
    return Cl.Item(ItemName,PriceList,BuyLimit,ParameterList)
    
    
def FindColumnLength(FileRb,SheetName):

    Sheet = FileRb.sheet_by_name(SheetName)
    Column = Sheet.col_values(0)
    for i in range(len(Column)):
        if(Column[i]==""):
            return i+1
    
    return len(Column)    


def TrueCondition(Item):
    return True

def DeleteSheet(FileWb,FileRb,SheetName):
    Names = FileRb.sheet_names()
    
    SheetFound = False
    index = 0
    for i in range(len(Names)):
        if(Names[i] == SheetName):
            index = i
            SheetFound = True
            break
            
    if(SheetFound):
        del FileWb._Workbook__worksheets[index]
    else:
        print("can't delete:'"+SheetName+"', sheet not found!")

def DoQuery(ParameterList,Condition=TrueCondition,FileName="Data",PriceSheetName="PriceSheet",ResultSheetName="ResultSheet"):
    
    #Finds file
    Path = MakePath(FileName,"xls")
    FileWb,FileRb = ReadExcel(Path)
    
    #Delete old resultsheet
    DeleteSheet(FileWb,FileRb,ResultSheetName)
    FileWb.save(Path)
    
    FileWb,FileRb = ReadExcel(Path)
    
    FileWb.set_colour_RGB(0x0A, 255, 80, 80)# Red
    FileWb.set_colour_RGB(0x11, 86, 160, 61)# Green
    
    #Create new empty Result sheet
    ResultSheet = Sa.FitSheetWrapper(FileWb.add_sheet(ResultSheetName,cell_overwrite_ok=True))
        
 
    #Create Header
    ResultSheet.write(0,0,"Item",Cl.StyleNormal)
    ResultSheet.write(0,1,"BuyLimit",Cl.StyleNormal) 
                
    for j in range(len(ParameterList)):
        ResultSheet.write(0,j+2,ParameterList[j].Name,Cl.StyleNormal)
        
    #Find vertical length of PriceSheet
    ColumnLength = FindColumnLength(FileRb,PriceSheetName)
        
    index = 0
    for i in range(1,ColumnLength):
        Item = CreateItem(i,ParameterList,FileRb,PriceSheetName)
            
            
        if(Condition(Item)):
            index +=1
        else:
            continue
                    
            
        ResultSheet.write(index,0,Item.Name,Cl.StyleNormal)
        ResultSheet.write(index,1,str(Item.BuyLimit),Cl.StyleNormal)
            
        for j in range(len(ParameterList)):
            Value =  Item.ValueList[j]
            Style = ParameterList[j].StyleFunc(Value)
                
            ResultSheet.write(index,j+2,str(Value),Style)
                
    FileWb.save(Path)
    print("Query completed!")


#-----Conditions-----#

def RisingCondition(Item):
    Cond1 = Item.ValueDict["CompToAverage"] <= 0
    Cond2 = Item.ValueDict["W1Slope"] > 0
    return Cond1 and Cond2

#-----Conditions-----#

#-----Main-----#






DoQuery([Cl.CurrentPrice,Cl.Average,Cl.CompToAverage,Cl.W1Slope,Cl.Periodicity,Cl.MaxProfit,Cl.M1Slope],RisingCondition,PriceSheetName="UnfilteredPriceSheet")

#-----Main-----#
    