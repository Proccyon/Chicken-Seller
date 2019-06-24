from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup
import numpy as np


#Gets the price of a single item

def MakeItemUrl(Id,Name):
    return"http://services.runescape.com/m=itemdb_rs/"+ Name + "viewitem?obj="+str(Id)

def RemoveNonsense(DataList):
    del DataList[0:19]
    NewDataList = []
    for Data in DataList:
        if("Date" in Data):
            NewDataList.append(Data)
    return NewDataList
    
def MakePriceList(DataList):
    DateList = []
    PriceList = []
    
    for data in DataList:
        DateList.append(data.split("'")[1])
        PriceList.append(int(data.split(",")[1]))
    return DateList,PriceList


def GetPriceList(Url):
    
    with closing(get(Url, stream=True)) as resp:
        html  = resp.content
        
    
    StringList = []
    html = BeautifulSoup(html, 'html.parser')
    for script in html.select('script'):
        s = script.text
        if("Date"in s):
            StringList.append(s)
    if(len(StringList)==0):
        return []
    
    RawData = StringList[0]
    RawDataList = RawData.split("\n")
    DataList = RemoveNonsense(RawDataList)
    DateList, PriceList = MakePriceList(DataList)
    del PriceList[0:120]
    return PriceList