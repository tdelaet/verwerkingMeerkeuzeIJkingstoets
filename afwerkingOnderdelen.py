import os
import pandas as pd
import sys
import shutil
import numpy
  
def kopieerQSF(jaar,toets):
    qsfSourceFilename= "../" + jaar + "_" +  toets + "_TOTAAL/printenscan/antwoorden.qsf"
    qsfTargetFilename= "../" + jaar + "_" +  toets + "/antwoorden.qsf"
    shutil.copyfile(qsfSourceFilename, qsfTargetFilename)
    
def genereerPuntenBestand(jaar,toets,sessie,onderdelen,regelFeedbackgroep):
    #lees punten van TOTAAL
    puntenFilename= "../" + jaar + "_" +  toets + "_TOTAAL/output/punten.xls"
    punten_onderdeel = pd.read_excel(puntenFilename)#,dtype=str)

    columns_punten = [punten_onderdeel.columns[x] for x in [0,1,3,4,5]]
    punten_compose=punten_onderdeel[columns_punten]
    namen_nieuw = ["nummer","TOTAAL","juist","fout","blanco"]
    punten_compose.columns=namen_nieuw
    
    
    #namen_nieuw = ["FeedbackGroep","ijkingstoetssessie","ijkID","Voornaam","Naam"]
    #df = pd.DataFrame(columns=namen_nieuw)
    #punten_compose[namen_nieuw] = df[namen_nieuw]

    punten_compose.insert(1,"ijkingstoetssessie",numpy.ones(punten_compose.shape[0]) * sessie)
    punten_compose.insert(1,"ijkID",[""]* punten_compose.shape[0])
    punten_compose.insert(0,"Voornaam",[""]* punten_compose.shape[0])
    punten_compose.insert(0,"Naam",[""]* punten_compose.shape[0])
    
   
    for onderdeel in onderdelen:
        onderdeelFolder = "../" + jaar + "_" +  toets + "_"+ onderdeel
        if not os.path.exists(onderdeelFolder):
             print("Error: folder " + onderdeelFolder + " does not exist")
             sys.exit()
        puntenFilename= onderdeelFolder + "/output/punten.xls"
        if not os.path.exists(puntenFilename):
             print("Error: file " + puntenFilename + " does not exist")
             sys.exit()
        punten_onderdeel = pd.read_excel(puntenFilename)
        columns_punten = [punten_onderdeel.columns[x] for x in [1,3,4,5]]
        punten_onderdeel_selected=punten_onderdeel[columns_punten]
        namen_nieuw = ["score" + onderdeel,"juist"+ onderdeel,"fout"+ onderdeel,"blanco"+ onderdeel]
        punten_onderdeel_selected.columns=namen_nieuw
        punten_compose[namen_nieuw]=punten_onderdeel_selected
 
    legeOnderdelen=[]
    for letter in range(97+len(onderdelen),97+8):
        legeOnderdelen.append(chr(letter).capitalize())

    for onderdeel in legeOnderdelen:
        namen_nieuw = ["score" + onderdeel,"juist"+ onderdeel,"fout"+ onderdeel,"blanco"+ onderdeel]
        df = pd.DataFrame(columns=namen_nieuw)
        punten_compose[namen_nieuw] = df[namen_nieuw]
        
    
    feedbackgroep=bepaalFeedbackGroep(punten_compose,regelFeedbackgroep)
    punten_compose.insert(5,"FeedbackGroep",feedbackgroep)
        
    punten_compose.to_excel("../" + jaar + "_" +  toets +"/resultaten.xlsx",sheet_name="punten",index=False)

def bepaalFeedbackGroep(df,regelFeedbackgroep):
    feedbackgroep = [""]* df.shape[0]
    feedbackgroepA = [False]* df.shape[0]
    feedbackgroepB = [False]* df.shape[0]
    feedbackgroepC = [False]* df.shape[0]
    feedbackgroepD = [False]* df.shape[0]
    feedbackgroepE = [False]* df.shape[0]
    feedbackgroepF = [False]* df.shape[0]
    
    if regelFeedbackgroep == "iedereenA":
        feedbackgroepA = [True]* df.shape[0]
    if regelFeedbackgroep == "geslaagdTotaal":            
        feedbackgroepA = (df["TOTAAL"].values>=10)
        feedbackgroepB = [not x for x in feedbackgroepA]
    if regelFeedbackgroep == "ia":            
        feedbackgroepA = (df["TOTAAL"].values>=10) & (df["scoreB"].values>=10)
        feedbackgroepB = [not x for x in feedbackgroepA]
    if regelFeedbackgroep == "dw":            
        feedbackgroepA = (df["TOTAAL"].values>=9)
        feedbackgroepB = (df["TOTAAL"].values<9) & (df["TOTAAL"].values>=5)
        feedbackgroepC = (df["TOTAAL"].values<5)

    feedbackgroep = numpy.where(feedbackgroepA,"A",feedbackgroep)
    feedbackgroep = numpy.where(feedbackgroepB,"B",feedbackgroep)
    feedbackgroep = numpy.where(feedbackgroepC,"C",feedbackgroep)
    feedbackgroep = numpy.where(feedbackgroepD,"D",feedbackgroep)
    feedbackgroep = numpy.where(feedbackgroepE,"E",feedbackgroep)
    feedbackgroep = numpy.where(feedbackgroepF,"F",feedbackgroep)
    #print(feedbackgroep)
    return feedbackgroep
    