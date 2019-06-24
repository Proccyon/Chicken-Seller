from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup
import matplotlib.pyplot as plt
import numpy as np
from GetPrice import GetPriceList
import xlwt

#-----GlobalVariables-----#
borders= xlwt.Borders()
borders.left= 1
borders.right= 1
borders.top= 1
borders.bottom= 1

StyleRed = xlwt.easyxf('pattern: pattern solid, fore_colour red')
StyleGreen = xlwt.easyxf('pattern: pattern solid, fore_colour green')
StyleBlue = xlwt.easyxf('pattern: pattern solid, fore_colour periwinkle')
StyleNormal = xlwt.easyxf('pattern: pattern solid, fore_colour white')

StyleRed.borders = borders
StyleGreen.borders = borders
StyleBlue.borders = borders
StyleNormal.borders = borders

#-----GlobalVariables-----#


#-----GlobalFunctions-----#


def TransformPrice(Price):
    if(np.abs(Price) >= 1E6):
        return str(round(Price/1E6,2))+"m"
    if(np.abs(Price) >= 1E3):
        return str(round(Price/1E3,2))+"k"  
    
    return str(round(Price,2))

def MakeItemUrl(Id,Name):
    return"http://services.runescape.com/m=itemdb_rs/"+ Name + "viewitem?obj="+str(Id)

#-----GlobalFunctions-----#


#-----ItemClass-----#
class Item:
    def __init__(self,Name,Url,BuyLimit):
        self.BuyLimit = BuyLimit
        self.Name = Name
        self.Url = Url
        self.Price = GetPriceList(Url)
        self.CurrentPrice = self.Price[-1]
        self.Rise = self.Price[-1]-self.Price[-2]
        self.D3Rise = self.Price[-1]-self.Price[-4]
        self.W1Rise = self.Price[-1]-self.Price[-8]
        self.M1Rise = self.Price[-1]-self.Price[-31]
        self.Average = int(round(np.average(self.Price),0))
        self.CompToAverage = int(round(self.CurrentPrice - self.Average,0))
        self.MaxProfit,self.MaxLoss = self.CalcMaxProfit(self.Price)
        self.Periodicity = round(np.amin([self.MaxProfit,self.MaxLoss]) / self.Average,2)
        self.Slope = round(self.LineFit(self.Price) / self.Average,2)
        self.CreateLists()
        
        
    def CalcMaxProfit(self,PriceList):
        MaxProfit = 0
        MaxLoss = 0
        for i in range(len(PriceList)-1):
            DeltaPrice = PriceList[i+1]-PriceList[i]
            if(DeltaPrice > 0):
                MaxProfit += DeltaPrice
            if(DeltaPrice < 0):
                MaxLoss -= DeltaPrice
                
        return MaxProfit,MaxLoss

    def LineFit(self,Price):
        t = range(len(Price))
        Parameters = np.polyfit(t,Price,1)
        return Parameters[0]*len(Price)
        
    def HighLow(Bool):
        if(Bool):
            return "Higher than average"
        return "Lower than average"
        
    def RG(self,Value):
        
        if(Value > 0):
            return StyleGreen
        if(Value < 0):
            return StyleRed
        return StyleBlue
            
    
    def CreateLists(self):
        self.NameList = ["Name","Price","Change[1d]","Average","Price-Average","MaxProfit","Periodicity","Slope",
        "BuyLimit"]
        self.ValueList = [self.Name,self.CurrentPrice,self.Rise,self.Average,self.CompToAverage,self.MaxProfit,
        self.Periodicity,self.Slope,self.BuyLimit]
        self.StyleList = [StyleNormal,StyleNormal,self.RG(self.Rise),StyleNormal,self.RG(self.CompToAverage),
        StyleNormal,StyleNormal,self.RG(self.Slope),StyleNormal]
        

#-----ItemClass-----#


#-----PurchaseClass-----#
class Purchase:
    def __init__(self,Item,Amount,BoughtPrice):
        self.Item = Item
        self.Amount = Amount
        self.BoughtPrice = BoughtPrice
        self.Profit = Amount*(Item.Price[-1]-BoughtPrice)

#-----PurchaseClass-----#




#-----ParameterClass-----#
#A parameter is a property of an item, for example it's 6-month average
#ParmFunc is a function that calculates the value of the property
class Parameter:
    def __init__(self,Name,ParmFunc):
        self.Name = Name
        self.ParmFunc = ParmFunc



#-----ParameterClass-----#


#Hello there

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


            
    
