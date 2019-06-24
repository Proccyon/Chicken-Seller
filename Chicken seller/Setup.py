from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup
import numpy as np
import xlrd
import xlwt
from xlutils.copy import copy
import sys,os
import time





#-----GlobalFunctions-----#
def GetSheet(FileRb,FileWb,TargetName):
    
    Names = FileRb.sheet_names()
    for i in range(len(Names)):
        Sheet = FileWb.get_sheet(i)
        if(Names[i] == TargetName):
            return Sheet
            
    return None

def MakePath(FileName,FileType):
    return os.path.dirname(sys.argv[0])+r'/ '[0] +str(FileName)+"."+FileType

def MakeItemUrl(Id,Name):
    return"http://services.runescape.com/m=itemdb_rs/"+ Name + "viewitem?obj="+str(Id)  

#-----GlobalFunctions-----#

def GetItemName(Id):

    Url = "https://www.runelocus.com/item-details/?item_id="+str(Id)
    Counter = 0
    
    while True:
        try:
            with closing(get(Url, stream=True)) as resp:
                RawHtml  = resp.content
                
            RawHtml = BeautifulSoup(RawHtml, 'html.parser')
            StringList = []
            for td in RawHtml.select('h2'):
                s = td.text
                StringList.append(s)
                
            return (StringList[0].split("'")[1]).replace(" ","_")
            break
            
        except:
            print("Getting item name failed. Retrying...")
            if(Counter > 2):
                return None
            Counter += 1
            time.sleep(1)
  
        
def CheckCorrectId(Id,Name):
    Url = MakeItemUrl(Id,Name)
    Counter = 0
    
    while(True):
        try:
            with closing(get(Url, stream=True)) as resp:
                RawHtml  = resp.content
                
            RawHtml = BeautifulSoup(RawHtml, 'html.parser')
            StringList = []
            for td in RawHtml.select('h3'):
                s = td.text
                StringList.append(s)
                
            return not ("Sorry" in StringList[0])
            break
            
        except:
            print("Id check failed. Retrying...")
            if(Counter > 3):
                return False
            Counter +=1
            time.sleep(8)


def ReadExcel(Path):
    rb = xlrd.open_workbook(Path)
    return copy(rb),rb

def FindLastId(FileRb,SheetName):

    IdSheet = FileRb.sheet_by_name(SheetName)
    Column = IdSheet.col_values(0)
    return int(Column[-1]),len(Column)

def MakeIdList(FileName="Data",IdStep=500,SheetName="IdSheet"):
    
    Path = MakePath(FileName,"xls")
    FileWb,FileRb = ReadExcel(Path)

    if(GetSheet(FileRb,FileWb,SheetName) == None):
        IdSheet = FileWb.add_sheet(SheetName)
        IdSheet.write(0,0,"ID")
        IdSheet.write(0,1,"Item Name")
        IdSheet.write(0,2,"Item Url")
        LastId = 1
        ColumnLength = 1
    else:
        IdSheet = GetSheet(FileRb,FileWb,SheetName)
        LastId,ColumnLength = FindLastId(FileRb,SheetName)
    
    Index = 1
    Counter = 0
    SaveCount = 1
    MaxId = LastId + IdStep
    
    for Id in range(LastId+1,MaxId):
        Name = GetItemName(Id)
        if(Name==None):
            continue

        c = False
        for i in range(5):
            if(CheckCorrectId(Id,Name)):
                c = True
                break
            time.sleep(0.5)
            
        if(c):
                
                print(Name+"["+str(Id)+"]")
                IdSheet.write(Index+ColumnLength-1,0,str(Id))
                IdSheet.write(Index+ColumnLength-1,1,Name)
                IdSheet.write(Index+ColumnLength-1,2,MakeItemUrl(Id,Name))
                Index += 1
                
                if(Counter >= SaveCount):
                    print("Saving...")
                    FileWb.save(Path)
                    Counter = 0

                else:
                    Counter += 1
                          
    FileWb.save(Path)
            

#-----Main-----#

MakeIdList(IdStep=2000)

#-----Main-----#
