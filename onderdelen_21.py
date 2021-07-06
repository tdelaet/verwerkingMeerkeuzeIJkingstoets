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
toets = "fa21"
editie= "juli "+ jaar

aantal_onderdelen = 4



onderdelen=[]
for letter in range(97,97+aantal_onderdelen):
    onderdelen.append(chr(letter))
    

correctAnswers = numpy.loadtxt("../" + jaar + "_" +  toets + "/onderdelen/sleutel_" + jaar+ "_"+ toets+ ".txt",delimiter=',',dtype="str")
OMRfilename = "../" + jaar + "_" +  toets + "/OMR/" + jaar+ "_"+ toets+ "_OMRoutput_all.xlsx"


OMR = pd.read_excel(OMRfilename,dtype=str)
OMR["vragenreeks"] = OMR["vragenreeks"].astype(str).astype(int)
#OMR.as_matrix()


outputFolder = "../" + jaar + "_" +  toets
if not os.path.exists(outputFolder):
    os.makedirs(outputFolder)
numpy.savetxt(outputFolder + "/sleutel_" + jaar+ "_"+ toets + ".txt",[correctAnswers],delimiter=',',fmt="%s")


for onderdeel in onderdelen:
    outputFolder = "../" + jaar + "_" +  toets + "_"+ onderdeel.capitalize()
    if not os.path.exists(outputFolder):
        os.makedirs(outputFolder)
    outputFolderOMR = "../" + jaar + "_" +  toets + "_"+ onderdeel.capitalize() + "/OMR"
    if not os.path.exists(outputFolderOMR):
        os.makedirs(outputFolderOMR)
   
    vragen_onderdeel = numpy.loadtxt("../" + jaar + "_" +  toets + "/onderdelen/" + jaar+ "_"+ toets+ "_"+ onderdeel.capitalize() + ".txt",delimiter=',',dtype="int")
    
    #get OMR with just questions of onderdeel
    namen_onderdeel = ["ijkID","vragenreeks"]

    counter=1
    namen_onderdeelNieuw= ["ijkID","vragenreeks"]
    for x in vragen_onderdeel:
        namen_onderdeel.append("Vraag"+str(x))
        namen_onderdeelNieuw.append("Vraag"+str(counter))
        counter=counter+1
    outputOMR= outputFolderOMR + "/" + jaar+ "_"+ toets+  "_"+ onderdeel.capitalize() +  "_OMRoutput_all.xlsx"
    OMR_onderdeel = OMR[namen_onderdeel]
    #rename questions to Vraag1,Vraag2, ...
    OMR_onderdeel.columns=namen_onderdeelNieuw
    
    OMR_onderdeel.to_excel(outputOMR,sheet_name="outputScan",index=False)
    

    # get sleutel with just questions of onderdeel
    if (vragen_onderdeel.ndim==0):
        correctAnswers_loc = correctAnswers[vragen_onderdeel]
    else:
        correctAnswers_loc = [correctAnswers[x-1] for x in vragen_onderdeel]
 
    # get overview of questions with just questions of onderdeel
    numpy.savetxt(outputFolder + "/sleutel_" + jaar+ "_"+ toets+ "_"+ onderdeel.capitalize() + ".txt",[correctAnswers_loc],delimiter=',',fmt="%s")
    numpy.savetxt(outputFolder + "/vragen_" + jaar+ "_"+ toets+ "_"+ onderdeel.capitalize() + ".txt",[vragen_onderdeel],delimiter=',',fmt="%i")
    
