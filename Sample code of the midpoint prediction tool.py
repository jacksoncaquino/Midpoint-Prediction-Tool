# -*- coding: utf-8 -*-
"""
Created on Fri Nov  4 08:47:56 2022

@author: Jackson Aquino
"""

def CalculateMidpointFromBestLevelComposite(CompMarket, JobFamily, ManagementLevel):
    tempdf = GetBenchmarkDataForJobFamily(CompMarket, JobFamily)
    if len(tempdf) == 0:
        return "nan"
    indice = tempdf["Incs"].idxmax()
    bestLevel = tempdf["Level"].loc[[indice]].values[0]
    bestLevelIncs = tempdf["Incs"].loc[[indice]].values[0]
    compositeBestLevel = tempdf["Market Composite"].loc[[indice]].values[0]
    global message
    message = message + "\n" + bestLevel + " has " + str(bestLevelIncs) + " incumbents in " + CompMarket + " with a composite of " + str(compositeBestLevel)
    print(bestLevel,"has",bestLevelIncs,"incumbents in",CompMarket,"with a composite of",compositeBestLevel)
    thisProgression = GetTypicalProgression(CompMarket,ManagementLevel,bestLevel)
    if thisProgression == "":
        return ""
    
    if thisProgression < 0:
        thisProgression = thisProgression * -1
        thisProgression = 1/(1+(thisProgression/100))
    else:
        thisProgression = 1+thisProgression/100
    
    indiceFrom = getLevelIndex(ManagementLevel)
    indiceTo = getLevelIndex(bestLevel)
    
    if (indiceFrom < indiceTo) or (indiceFrom == indiceTo and (bestLevel[0] == "M" or bestLevel[0]=="E")):
        if thisProgression > 1:
            thisProgression = 1/thisProgression
    else:
        if thisProgression < 1:
            thisProgression = 1/thisProgression
    
    return smartround(float(compositeBestLevel) * float(thisProgression))