from openpyxl import load_workbook, Workbook
from difflib import SequenceMatcher
import pandas

"""Contains all the common functions used by cleaning codes
        """

# Function which calculates the similarity percentage between two strings
def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

# Function which finds the most similar match for each family
def familyMatch(fam, families):
    # Secial Case for ICT
    rankDict = {}
    if 'ICT' in fam:
        score = 0.8
        rankDict['ICT'] = score
    else:
        # Checks if fam appears in families or similar
        for each in families:
            score = similar(fam, each)
            rankDict[each] = score
    sortedRank = sorted(rankDict.items(), key=lambda x:x[1], reverse=True)
    sortedRankDict = dict(sortedRank)
    # Return the highest scored item (Tuple containing both name and score)
    bestGuess = sortedRank[0]
    return bestGuess 

# Function which finds the most similar match for each product
def productMatch(prod, GARprod):
    # Checks if product appears in GAR product register or similar
    rankDict = {}
    for each in GARprod:
        if each.value == None:
            each.value = 'None'
        score = similar(prod, each.value)
        rankDict[each.value] = score
    sortedRank = sorted(rankDict.items(), key=lambda x:x[1], reverse=True)
    sortedRankDict = dict(sortedRank)
    # print(sortedRankDict)
    return sortedRank[0]