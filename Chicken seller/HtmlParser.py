from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup
import matplotlib.pyplot as plt
import numpy as np
from GetPrice import GetPriceList

#-----GlobalFunctions-----#
def LineFit(Price):
    t = range(len(Price))
    Parameters = np.polyfit(t,Price,1)
    return Parameters[0]*len(Price)

def TransformPrice(Price):
    if(np.abs(Price) >= 1E6):
        return str(round(Price/1E6,2))+"m"
    if(np.abs(Price) >= 1E3):
        return str(round(Price/1E3,2))+"k"  
    
    return str(round(Price,2))

def HighLow(Bool):
    if(Bool):
        return "Higher than average"
    return "Lower than average"


#-----GlobalFunctions-----#

class Item:
    def __init__(self,Name,Url):
        self.Name = Name
        self.Url = Url
        self.Price = GetPriceList(Url)
        self.CurrentPrice = self.Price[-1]
        self.Rise = self.Price[-1]-self.Price[-2]
        self.Average = np.average(self.Price)
        self.CompToAverage = self.CurrentPrice - self.Average
        self.MaxProfit,self.MaxLoss = CalcMaxProfit(self.Price)
        self.Periodicity = np.amin([self.MaxProfit,self.MaxLoss]) / self.Average
        self.Slope = LineFit(self.Price) / self.Average

class Purchase:
    def __init__(self,Item,Amount,BoughtPrice):
        self.Item = Item
        self.Amount = Amount
        self.BoughtPrice = BoughtPrice
        self.Profit = Amount*(Item.Price[-1]-BoughtPrice)


def PrintItemInfo(ItemList):
    print("Item | Price | Rise | Avergage | CompAverage | MaxProfit | Periodicity | Slope")
    for Item in ItemList:       
        print(Item.Name+" | "+TransformPrice(Item.CurrentPrice)+" | "+TransformPrice(Item.Rise)
        +" | "+TransformPrice(Item.Average)+" | "+TransformPrice(Item.CompToAverage)
        +" | "+TransformPrice(Item.MaxProfit)+" | "+str(round(Item.Periodicity,2))
        +" | "+str(round(Item.Slope,2))
        )

def PrintProfitInfo(PurchaseList):
    print("Item | Amount | BoughtPrice | Profit | Rise")
    for Purchase in PurchaseList:
        print(Purchase.Item.Name+" | "+str(Purchase.Amount)+" | "+TransformPrice(Purchase.BoughtPrice)+" | "+TransformPrice(Purchase.Profit))

def GetItemName(Id):

    Url = "https://www.runelocus.com/item-details/?item_id="+str(Id)
    with closing(get(Url, stream=True)) as resp:
        RawHtml  = resp.content
        
    RawHtml = BeautifulSoup(RawHtml, 'html.parser')
    StringList = []
    for td in RawHtml.select('h2'):
        s = td.text
        StringList.append(s)
        
    return (StringList[0].split("'")[1]).replace(" ","_")
    
def MakeItemUrl(Id,Name):
    return"http://services.runescape.com/m=itemdb_rs/"+ Name + "viewitem?obj="+str(Id)
    
def CheckCorrectId(Id,Name):
    Url = MakeItemUrl(Id,Name)
    
    with closing(get(Url, stream=True)) as resp:
        RawHtml  = resp.content
        
    RawHtml = BeautifulSoup(RawHtml, 'html.parser')
    StringList = []
    for td in RawHtml.select('h3'):
        s = td.text
        StringList.append(s)
        
    return not ("Sorry" in StringList[0])

def MakeItemList(IdStart,IdEnd):

    ItemList = []
    for Id in range(IdStart,IdEnd):
        
        Name = GetItemName(Id)
        
        if(CheckCorrectId(Id,Name)):
            ItemList.append(Item(Name,MakeItemUrl(Id,Name)))
            
            
    return ItemList
            
def CalcMaxProfit(PriceList):
    MaxProfit = 0
    MaxLoss = 0
    for i in range(len(PriceList)-1):
        DeltaPrice = PriceList[i+1]-PriceList[i]
        if(DeltaPrice > 0):
            MaxProfit += DeltaPrice
        if(DeltaPrice < 0):
            MaxLoss -= DeltaPrice
            
    return MaxProfit,MaxLoss
            
    
ItemList = MakeItemList(350,370)
PrintItemInfo(ItemList)
            





#ArmaBoots = Item("ArmaBoots","http://services.runescape.com/m=itemdb_rs/Armadyl_boots/viewitem?obj=25010")
#ArmaBuckler = Item("ArmaBuckler","http://services.runescape.com/m=itemdb_rs/Armadyl_buckler/viewitem?obj=25013")
#ArmaGloves = Item("ArmaGloves","http://services.runescape.com/m=itemdb_rs/Armadyl_gloves/viewitem?obj=25016")
#SoftClay = Item("SoftClay","http://services.runescape.com/m=itemdb_rs/Soft_clay/viewitem?obj=1761")
#RawBird = Item("Rawbird","http://services.runescape.com/m=itemdb_rs/Raw_bird_meat/viewitem?obj=9978")
#BloodNecklace = Item("BloodNecklace","http://services.runescape.com/m=itemdb_rs/Blood_necklace_shard/viewitem?obj=32692")
#Cinderbane = Item("Cinderbane","http://services.runescape.com/m=itemdb_rs/Cinderbane_gloves/viewitem?obj=41106")
#Bond = Item("Bond","http://services.runescape.com/m=itemdb_rs/Bond/viewitem?obj=29492")
#Zaryte = Item("Zaryte","http://services.runescape.com/m=itemdb_rs/Zaryte_bow/viewitem?obj=20171")
#SubGown = Item("SubGown","http://services.runescape.com/m=itemdb_rs/Gown_of_subjugation/viewitem?obj=24998")
#SubGarb = Item("SubGarb","http://services.runescape.com/m=itemdb_rs/Garb_of_subjugation/viewitem?obj=24995")
#ArmaStaff = Item("ArmaStaff","http://services.runescape.com/m=itemdb_rs/Armadyl_battlestaff/viewitem?obj=21777")
#AirRune = Item("AirRune","http://services.runescape.com/m=itemdb_rs/Air_rune/viewitem?obj=556")
#FireRune = Item("FireRune","http://services.runescape.com/m=itemdb_rs/Fire_rune/viewitem?obj=554")
#ItemList = [ArmaBoots,ArmaBuckler,ArmaGloves,SoftClay,RawBird,BloodNecklace,Cinderbane,Bond,Zaryte,
#SubGown,SubGarb,ArmaStaff,AirRune,FireRune]
#BoughtList = []

PrintItemInfo(ItemList)
            
    
