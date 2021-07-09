import os
import pandas as pd
import sys
import shutil

def kopieerQSF(jaar,toets):
    qsfSourceFilename= "../" + jaar + "_" +  toets + "_TOTAAL/printenscan/antwoorden.qsf"
    qsfTargetFilename= "../" + jaar + "_" +  toets + "/antwoorden.qsf"
    shutil.copyfile(qsfSourceFilename, qsfTargetFilename)
    
def genereerPuntenBestand(jaar,toets,onderdelen):
    #lees punten van TOTAAL
    puntenFilename= "../" + jaar + "_" +  toets + "_TOTAAL/output/punten_geheel.xls"
    punten_onderdeel = pd.read_excel(puntenFilename)#,dtype=str)

    columns_punten = [punten_onderdeel.columns[x] for x in [0,1,3,4,5]]
    punten_compose=punten_onderdeel[columns_punten]
    namen_nieuw = ["nummer","TOTAAL","juist","fout","blanco"]
    punten_compose.columns=namen_nieuw

    
    for onderdeel in onderdelen:
        onderdeelFolder = "../" + jaar + "_" +  toets + "_"+ onderdeel
        if not os.path.exists(onderdeelFolder):
             print("Error: folder " + onderdeelFolder + " does not exist")
             sys.exit()
        puntenFilename= onderdeelFolder + "/output/punten_geheel.xls"
        if not os.path.exists(puntenFilename):
             print("Error: file " + puntenFilename + " does not exist")
             sys.exit()
        punten_onderdeel = pd.read_excel(puntenFilename)
        columns_punten = [punten_onderdeel.columns[x] for x in [1,3,4,5]]
        punten_onderdeel_selected=punten_onderdeel[columns_punten]
        namen_nieuw = ["score" + onderdeel.capitalize(),"juist"+ onderdeel.capitalize(),"fout"+ onderdeel.capitalize(),"blanco"+ onderdeel.capitalize()]
        punten_onderdeel_selected.columns=namen_nieuw
        punten_compose[namen_nieuw]=punten_onderdeel_selected
 

    punten_compose.to_excel("../" + jaar + "_" +  toets +"/resultaten.xlsx",sheet_name="punten",index=False)
