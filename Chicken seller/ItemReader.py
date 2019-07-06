import numpy as np
import time
import sys,os
import xlrd
import xlwt
from xlutils.copy import copy

import Classes as Cl
import SizeAdjuster as Sa
from GetPrice import GetPriceList

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

def MakeItemList(FileName="Data",SheetName="UnfilteredPriceSheet",IdSheetName="IdSheet"):
    
    #Finds file
    Path = MakePath(FileName,"xls")
    FileWb,FileRb = ReadExcel(Path)
    NameList, UrlList,BuyLimitList = GetUrlList(FileName,IdSheetName)


    #Finds the sheet or create new one if empty
    if(GetSheet(FileRb,FileWb,SheetName) == None):
        PriceSheet = Sa.FitSheetWrapper(FileWb.add_sheet(SheetName))
        ColumnLength = 1
    else:
        PriceSheet = Sa.FitSheetWrapper(GetSheet(FileRb,FileWb,SheetName))
        ColumnLength = FindColumnLength(FileRb,SheetName)
    
    
    #Creates header
    PriceSheet.write(0,0,"Item")
    PriceSheet.write(0,1,"BuyLimit")    
    for j in range(177):
        PriceSheet.write(0,j+2,"Price d"+str(j))
        
    Counter = 0
    #Loops through the IdSheet to create Item objects
    for i in range(ColumnLength,len(UrlList)):
        
        
        #If the BuyLimit cell is empt, assume it is 1
        if(BuyLimitList[i] == ""):
            BuyLimit = 1 
                    
        #If BuyLimit isn't empty, retrieve it from cell
        else:
            BuyLimit = int(float(str(BuyLimitList[i]).replace(",","")))
        

        #while loop is used to bypass host closing connection
        while(True):
                
            PriceList = GetPriceList(str(UrlList[i]))
            
            #Pricelist is empty list if host closes connection, keep trying till it works
            if(PriceList == []):
                print("Item creation failed. Retrying...")
                time.sleep(3) 
            else:
                break                
                                                    
        
        print("{} [{}/{}]".format(NameList[i],i,len(UrlList)))
    
        PriceSheet.write(i,0,NameList[i],Cl.StyleNormal)
        PriceSheet.write(i,1,str(BuyLimit),Cl.StyleNormal)
        
        for k in range(len(PriceList)):
            PriceSheet.write(i,k+2,str(PriceList[k]),Cl.StyleNormal)
    
        #Save after every 10 items
        if(Counter >= 10):
            print("saving...")
            FileWb.save(Path)
            Counter = 0
        else:
            Counter +=1
    
    FileWb.save(Path)
    
    
    
    
#-----Main-----#

#ParameterList = [Cl.CurrentPrice,Cl.W1Slope,Cl.Average,Cl.CompToAverage,Cl.MaxProfit,Cl.Periodicity]

MakeItemList()

#-----Main-----#