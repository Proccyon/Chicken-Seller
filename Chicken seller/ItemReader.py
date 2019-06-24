import numpy as np
import sys,os
import Classes as Cl
import xlrd
import xlwt
from xlutils.copy import copy
import Classes as Cl
import time
import SizeAdjuster as Sa


#Reads the price of all items in the filtered id list and puts it into excel sheet




def GetSheet(FileRb,FileWb,TargetName):
    
    Names = FileRb.sheet_names()
    for i in range(len(Names)):
        Sheet = FileWb.get_sheet(i)
        if(Names[i] == TargetName):
            return Sheet    
    return None

def ReadExcel(Path):
    rb = xlrd.open_workbook(Path,formatting_info=True)
    return copy(rb),rb 

def MakePath(FileName,FileType):
    return os.path.dirname(sys.argv[0])+r'/ '[0] +str(FileName)+"."+FileType

def GetUrlList(FileName,SheetName):
    Path = MakePath(FileName,"xls")
    FileWb,FileRb = ReadExcel(Path)
    IdSheet = FileRb.sheet_by_name(SheetName)
    return IdSheet.col_values(1),IdSheet.col_values(2),IdSheet.col_values(3)
    
def FindColumnLength(FileRb,SheetName):

    IdSheet = FileRb.sheet_by_name(SheetName)
    Column = IdSheet.col_values(0)
    for i in range(len(Column)):
        if(Column[i]==""):
            return i+1
    
    return len(Column)    

def MakeItemList(FileName="Data",SheetName="ItemSheet2",IdSheetName="FilteredIdSheet"):
    
    #Finds file
    Path = MakePath(FileName,"xls")
    FileWb,FileRb = ReadExcel(Path)
    NameList, UrlList,BuyLimitList = GetUrlList(FileName,IdSheetName)

    #Change colours to a lighter version
    FileWb.set_colour_RGB(0x0A, 255, 80, 80)# Red
    FileWb.set_colour_RGB(0x11, 86, 160, 61)# Green

    #Finds the sheet or create new one if empty
    if(GetSheet(FileRb,FileWb,SheetName) == None):
        ItemSheet = Sa.FitSheetWrapper(FileWb.add_sheet(SheetName))
        ColumnLength = 1
    else:
        ItemSheet = Sa.FitSheetWrapper(GetSheet(FileRb,FileWb,SheetName))
        ColumnLength = FindColumnLength(FileRb,SheetName)
    
    #Creates header        
    Item = Cl.Item(NameList[1],str(UrlList[1]),1)  
    for j in range(len(Item.NameList)):
        ItemSheet.write(0,j,Item.NameList[j])
        
    Counter = 0
    
    #Loops through the items
    for i in range(ColumnLength,len(UrlList)):
        
        #Initialize item
        while(True):
            try:
                if(BuyLimitList[i] == ""):
                    BuyLimit = 1
                else:
                    BuyLimit = int(float(str(BuyLimitList[i]).replace(",","")))
                    
                Item = Cl.Item(NameList[i],str(UrlList[i]),BuyLimit)
                #time.sleep(4)
                break
            except:
                print("Item creation failed. Retrying...")
                time.sleep(3) 
        
        print(Item.Name+"["+str(i)+"/"+str(len(UrlList))+"]")
        
        #Loop through item variables
        for k in range(len(Item.ValueList)):
            ItemSheet.write(i,k,str(Item.ValueList[k]),Item.StyleList[k])
    
        #Save after every 10 items
        if(Counter >= 10):
            print("saving...")
            FileWb.save(Path)
            Counter = 0
        else:
            Counter +=1
    
    FileWb.save(Path)
    
    
    
    
#-----Main-----#

MakeItemList()

#-----Main-----#