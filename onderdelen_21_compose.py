# -*- coding: utf-8 -*-
"""
Created on Mon Mar 31 13:58:48 2014

@author: tdelaet

Dit neemt aan dat de gebruikte sheet van excel file de volgende kolommen heeft (met eerste rij de naam van de kolom):
- studentennummer
- vragenreeks
- Vraag1, Vraag2, ... 
 en dit voor alle vragen (komt overeen met numQuestions)
"""

from xlrd import open_workbook
import string
import numpy
import matplotlib.pyplot as plt
from xlwt import Workbook
import matplotlib
import os
import pandas as pd

jaar = "2021"
toets = "hw21"
editie= "juli "+ jaar

aantal_onderdelen =6



onderdelen=[]
for letter in range(97,97+aantal_onderdelen):
    onderdelen.append(chr(letter))
    

outputFolder = "../" + jaar + "_" +  toets + "/output_goedgekeurd"
if not os.path.exists(outputFolder):
    os.makedirs(outputFolder)

#geheel
puntenFilename= "../" + jaar + "_" +  toets + "/output/punten_geheel.xls"
punten_onderdeel = pd.read_excel(puntenFilename)#,dtype=str)
#OMR["vragenreeks"] = OMR["vragenreeks"].astype(str).astype(int)

columns_punten = [punten_onderdeel.columns[x] for x in [0,1,3,4,5]]
punten_compose=punten_onderdeel[columns_punten]
namen_nieuw = ["nummer","TOTAAL","juist","fout","blanco"]
punten_compose.columns=namen_nieuw

    
for onderdeel in onderdelen:
    onderdeelFolder = "../" + jaar + "_" +  toets + "_"+ onderdeel.capitalize()
    if not os.path.exists(outputFolder):
         print("Error: folder " + onderdeelFolder + " does not exist")
         
    puntenFilename= onderdeelFolder + "/output/punten_geheel.xls"
    punten_onderdeel = pd.read_excel(puntenFilename)#,dtype=str)

    columns_punten = [punten_onderdeel.columns[x] for x in [1,3,4,5]]
    punten_onderdeel_selected=punten_onderdeel[columns_punten]
    namen_nieuw = ["score" + onderdeel.capitalize(),"juist"+ onderdeel.capitalize(),"fout"+ onderdeel.capitalize(),"blanco"+ onderdeel.capitalize()]
    punten_onderdeel_selected.columns=namen_nieuw
    #punten_compose= punten_compose.append(punten_onderdeel_selected, ignore_index=False)
    punten_compose[namen_nieuw]=punten_onderdeel_selected
 

punten_compose.to_excel(outputFolder+"/resultaten.xlsx",sheet_name="punten",index=False)