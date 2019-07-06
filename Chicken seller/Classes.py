from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup
import matplotlib.pyplot as plt
import numpy as np
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
    def __init__(self,Name,PriceList,BuyLimit,ParameterList):
        self.Name = Name
        self.BuyLimit = int(BuyLimit)
        self.PriceList = PriceList
        self.ParameterList = ParameterList

        self.CreateLists() #Unpacks information in ParameterList
        
        
    def HighLow(Bool):
        if(Bool):
            return "Higher than average"
        return "Lower than average"
        
            
    #Unpacks the information in ParameterList
    def CreateLists(self):
        
        self.NameList = []
        self.ValueList = []
        self.StyleList = []
        self.ValueDict = {}
        
        for Parameter in self.ParameterList:
            self.NameList.append(Parameter.Name)
            
            Value = Parameter.ParmFunc(self.PriceList,self.BuyLimit)
            self.ValueList.append(Value)
            self.ValueDict[Parameter.Name] = Value
            
            Style = Parameter.StyleFunc(Value)
            self.StyleList.append(Style)
            
#-----ItemClass-----#


#-----PurchaseClass-----#
class Purchase:
    def __init__(self,Item,Amount,BoughtPrice):
        self.Item = Item
        self.Amount = Amount
        self.BoughtPrice = BoughtPrice
        self.Profit = Amount*(Item.PriceList[-1]-BoughtPrice)

#-----PurchaseClass-----#




#-----ParameterClass-----#

#A parameter is a property of an item, for example it's 6-month average
#ParmFunc is a function that calculates the value of the property
#StyleFunc is a function that determines the excel style of the cell the value is in (red or green etc.)


class Parameter:
    def __init__(self,Name,ParmFunc,StyleFunc):
        self.Name = Name
        self.ParmFunc = ParmFunc
        self.StyleFunc = StyleFunc


#-----ParameterClass-----#



#-----StyleFunctions-----#
#Examples of stylefunctions

#Returns a green style when value > 0, returns a red style when value < 0
def RedGreenStyle(Value):
        
    if(Value > 0):
        return StyleGreen
    if(Value < 0):
        return StyleRed
    return StyleBlue

#Returns the default excel style with borders
def NormalStyle(Value):
    return StyleNormal


#-----StyleFunctions-----#

#-----ParameterExamples-----#

#Some example parameters that can be used
#Each parameter is equivalent to a row in the 'ItemSheet' sheet


#Creates a function that gives the price rise of an item over "Period" amount of days
def CreateRiseFunction(Period):
    return lambda PriceList,BuyLimit : PriceList[-1] - PriceList[-1-Period]

#Creates a Parameter that gives the price rise of an item over "Period" amount of days
def CreateRiseParameter(Name,Period,Style):
    return Parameter(Name,CreateRiseFunction(Period),Style)

#Does a linefit over "Period" amount of days and returns the slope times a constant
def LineFit(PriceList,Period):
    t = range(len(PriceList[(-1-Period):]))
    Parameters = np.polyfit(t,PriceList[(-1-Period):],1)
    return Parameters[0]*len(PriceList)   #len(price) is the amount of days in 6 months(constant)


#Creates a function that gives the price rise of an item over "Period" amount of days
def CreateSlopeFunction(Period):
    return lambda PriceList,BuyLimit : round(LineFit(PriceList,Period) / GetAverage(PriceList),2)


#Creates a Parameter that gives slope of an item over "Period" amount of days
def CreateSlopeParameter(Name,Period,Style):
    return Parameter(Name,CreateSlopeFunction(Period),Style)


#Function that finds the current price of an item
def GetCurrentPrice(PriceList,BuyLimit=1):
    return PriceList[-1]  
      
#Function that calculates the average price over 6 months  
def GetAverage(PriceList,BuyLimit=1):
    return int(round(np.average(PriceList),0))

#Function that calculates the difference between the current price and average price
def GetCompToAverage(PriceList,BuyLimit=1):
    return int(round(GetCurrentPrice(PriceList) - GetAverage(PriceList),0))

#Function that finds the maximum profit that could be made
def GetMaxProfit(PriceList,BuyLimit):
    InvestedItems = 0
    InvestedMoney = 0
    MaxProfit = 0
    
    for i in range(len(PriceList)):
        if(PriceList[i] < np.amax(PriceList[i:])-0.1):
            InvestedItems += 6*BuyLimit
            InvestedMoney += 6*BuyLimit*PriceList[i]
            
        else:
            MaxProfit += InvestedItems * PriceList[i] - InvestedMoney
            InvestedItems = 0
            InvestedMoney = 0
            
    return MaxProfit
    
    
def GetPeriodicity(PriceList,BuyLimit=1):
        TotalUp = 0
        TotalDown = 0
        for i in range(len(PriceList)-1):
            DeltaPrice = PriceList[i+1]-PriceList[i]
            if(DeltaPrice > 0):
                TotalUp += DeltaPrice
            if(DeltaPrice < 0):
                TotalDown -= DeltaPrice
                
        return round(np.amin([TotalUp,TotalDown]) / GetAverage(PriceList),2)


CurrentPrice = Parameter("CurrentPrice",GetCurrentPrice,NormalStyle)
Average = Parameter("Average",GetAverage,NormalStyle)
CompToAverage = Parameter("CompToAverage",GetCompToAverage,RedGreenStyle)
MaxProfit = Parameter("MaxProfit",GetMaxProfit,NormalStyle)
Periodicity = Parameter("Periodicity",GetPeriodicity,NormalStyle)
D1Rise = CreateRiseParameter("D1Rise",1,RedGreenStyle)
D3Rise = CreateRiseParameter("D3Rise",3,RedGreenStyle)
W1Rise = CreateRiseParameter("W1Rise",7,RedGreenStyle)
M1Rise = CreateRiseParameter("M1Rise",30,RedGreenStyle)
D1Slope = CreateSlopeParameter("D1Slope",1,RedGreenStyle)
D3Slope = CreateSlopeParameter("D3Slope",3,RedGreenStyle)
W1Slope = CreateSlopeParameter("W1Slope",7,RedGreenStyle)
M1Slope = CreateSlopeParameter("M1Slope",30,RedGreenStyle)

#-----ParameterExamples-----#



            
    
