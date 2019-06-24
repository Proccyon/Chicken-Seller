# -*- coding: utf-8 -*-
from HtmlParser import GetPriceList
import numpy as np

class Item:
    def __init__(self,Name,Url):
        self.Name = Name
        self.Url = Url
        self.Price = GetPriceList(Url)
        self.CurrentPrice = TransformPrice(self.Price[-1])
        self.Rise = TransformPrice(self.Price[-1]-self.Price[-2])
        self.AboveAverage = HighLow(self.Price[-1]>np.average(self.Price))

class Purchase:
    def __init__(self,Item,Amount,BoughtPrice):
        self.Item = Item
        self.Amount = Amount
        self.BoughtPrice = BoughtPrice
        self.Profit = Amount*(Item.Price[-1]-BoughtPrice)


def TransformPrice(Price):
    if(np.abs(Price) >= 1E6):
        return str(round(Price/1E6,2))+"m"
    if(np.abs(Price) >= 1E3):
        return str(round(Price/1E3,2))+"k"  
    
    return str(Price)

def HighLow(Bool):
    if(Bool):
        return "Higher than average"
    return "Lower than average"


def PrintItemInfo(ItemList):
    print("Item | Price | Rise | High/low")
    for Item in ItemList:       
        print(Item.Name+" | "+Item.CurrentPrice+" | "+Item.Rise+" | "+Item.AboveAverage)

def PrintProfitInfo(PurchaseList):
    print("Item | Amount | BoughtPrice | Profit | Rise")
    for Purchase in PurchaseList:
        print(Purchase.Item.Name+" | "+str(Purchase.Amount)+" | "+TransformPrice(Purchase.BoughtPrice)+" | "+TransformPrice(Purchase.Profit))

ArmaBoots = Item("ArmaBoots","http://services.runescape.com/m=itemdb_rs/Armadyl_boots/viewitem?obj=25010")
ArmaBuckler = Item("ArmaBuckler","http://services.runescape.com/m=itemdb_rs/Armadyl_buckler/viewitem?obj=25013")
ArmaGloves = Item("ArmaGloves","http://services.runescape.com/m=itemdb_rs/Armadyl_gloves/viewitem?obj=25016")
SoftClay = Item("SoftClay","http://services.runescape.com/m=itemdb_rs/Soft_clay/viewitem?obj=1761")
RawBird = Item("Rawbird","http://services.runescape.com/m=itemdb_rs/Raw_bird_meat/viewitem?obj=9978")
BloodNecklace = Item("BloodNecklace","http://services.runescape.com/m=itemdb_rs/Blood_necklace_shard/viewitem?obj=32692")
Cinderbane = Item("Cinderbane","http://services.runescape.com/m=itemdb_rs/Cinderbane_gloves/viewitem?obj=41106")
Bond = Item("Bond","http://services.runescape.com/m=itemdb_rs/Bond/viewitem?obj=29492")
Zaryte = Item("Zaryte","http://services.runescape.com/m=itemdb_rs/Zaryte_bow/viewitem?obj=20171")

ItemList = [ArmaBoots,ArmaBuckler,ArmaGloves,SoftClay,RawBird,BloodNecklace,Cinderbane,Bond]

BoughtList = []

PrintItemInfo(ItemList)