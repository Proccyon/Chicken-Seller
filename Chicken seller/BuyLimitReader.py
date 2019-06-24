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

def MakeWikiUrl(ItemName):
    return "https://runescape.fandom.com/wiki/Exchange:"+ItemName
    
def ReadExcel(Path):
    rb = xlrd.open_workbook(Path,formatting_info=True)
    return copy(rb),rb
    
def GetSheet(FileRb,FileWb,TargetName):
    
    Names = FileRb.sheet_names()
    for i in range(len(Names)):
        Sheet = FileWb.get_sheet(i)
        if(Names[i] == TargetName):
            return Sheet 

def MakePath(FileName,FileType):
    return os.path.dirname(sys.argv[0])+r'/ '[0] +str(FileName)+"."+FileType
#-----GlobalFunctions-----#


def FindColumnLength(FileRb,SheetName,N):

    IdSheet = FileRb.sheet_by_name(SheetName)
    Column = IdSheet.col_values(N)
    return len(Column)
    
def FindBlLength(FileRb,SheetName,MaxLength):
    IdSheet = FileRb.sheet_by_name(SheetName)
    BlColumn = IdSheet.col_values(3)
    for i in range(len(BlColumn)):
        if(BlColumn[i]==""):
            return i
    return MaxLength

def AddBuyLimit(FileName="Data",SheetName="IdSheet"):
    
    Path = MakePath(FileName,"xls")
    FileWb,FileRb = ReadExcel(Path)
    
    IdLength = FindColumnLength(FileRb,SheetName,0)
    BlLength = FindBlLength(FileRb,SheetName,IdLength) #Buy limit column length
    
    IdSheet = GetSheet(FileRb,FileWb,SheetName)
    IdSheetRb = FileRb.sheet_by_name(SheetName)
    IdSheet.write(0,3,"BuyLimit")
    
    for i in range(BlLength,IdLength):
        ItemName = IdSheetRb.col_values(1)[i]
        WikiUrl = MakeWikiUrl(ItemName)
        BuyLimit = GetBuyLimit(WikiUrl)
        if(BuyLimit==None):
            continue
        
        IdSheet.write(i,3,str(BuyLimit))
        #print(ItemName)
        
    FileWb.save(Path)
        
    
    

def GetBuyLimit(WikiUrl):
    
    Counter = 0
    while True:
        try:
            
            with closing(get(WikiUrl, stream=True)) as resp:
                RawHtml  = resp.content
                        
            RawHtml = BeautifulSoup(RawHtml, 'html.parser')
            StringList = []
            for td in RawHtml.select('li'):
                s = td.text
                StringList.append(s)
                        
            return (SearchListForString(StringList,"Exchange limit")).split()[2]
            
        except:
            print("Getting buy limit failed. Retrying...")
            time.sleep(0.5)
            if(Counter > 5):
                return None
            Counter += 1
    

def SearchListForString(List,TargetString):
    
    for String in List:
        if(TargetString in String):
            return String
    

#-----Main-----#

AddBuyLimit()
#-----Main-----#