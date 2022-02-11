import re

def extractTitle(listitem):
    listitem = str(listitem)
    pattern = r"'(.*?[^\\])'"
    targetString = str(re.findall(pattern,listitem))
    targetString1 = targetString.replace("[","")
    targetString2 = targetString1.replace("]","")
    nameOfWindow = targetString2.replace("'","")
    print(nameOfWindow)
