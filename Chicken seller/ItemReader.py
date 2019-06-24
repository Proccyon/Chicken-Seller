import numpy as np
import time
import sys,os
import xlrd
import xlwt
from xlutils.copy import copy

import Classes as Cl
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

def MakeItemList(ParameterList,FileName="Data",SheetName="ItemSheet3",IdSheetName="FilteredIdSheet"):
    
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
    ItemSheet.write(0,0,"Item")    
    for j in range(len(ParameterList)):
        ItemSheet.write(0,j+1,ParameterList[j].Name)
        
    Counter = 0
    #Loops through the IdSheet to create Item objects
    for i in range(ColumnLength,len(UrlList)):
        
        #while loop is used to bypass host closing connection
        while(True):
            try:
                #If the BuyLimit cell is empt, assume it is 1
                if(BuyLimitList[i] == ""):
                    BuyLimit = 1 
                    
                #If BuyLimit isn't empty, retrieve it from cell
                else:
                    BuyLimit = int(float(str(BuyLimitList[i]).replace(",","")))
                    
                
                Item = Cl.Item(NameList[i],str(UrlList[i]),BuyLimit,ParameterList)
                #time.sleep(4)
                break
                
            #If Host closes connection, try again after 3 seconds
            except:
                print("Item creation failed. Retrying...")
                time.sleep(3) 
        
        print(Item.Name+"["+str(i)+"/"+str(len(UrlList))+"]")
        
        #Loop through item Parameters and writes down the values in excel file
        ItemSheet.write(i,0,Item.Name,Cl.StyleNormal)
        for k in range(len(Item.ValueList)):
            ItemSheet.write(i,k+1,str(Item.ValueList[k]),Item.StyleList[k])
    
        #Save after every 10 items
        if(Counter >= 10):
            print("saving...")
            FileWb.save(Path)
            Counter = 0
        else:
            Counter +=1
    
    FileWb.save(Path)
    
    
    
    
#-----Main-----#

ParameterList = [Cl.CurrentPrice,Cl.W1Slope,Cl.Average,Cl.CompToAverage,Cl.MaxProfit,Cl.Periodicity]

MakeItemList(ParameterList)

#-----Main-----#