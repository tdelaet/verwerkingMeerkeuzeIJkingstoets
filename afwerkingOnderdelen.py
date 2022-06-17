import os
import pandas as pd
import sys
import shutil
import numpy
  
def kopieerQSF(jaar,toets):
    qsfSourceFilename= "../" + jaar + "_" +  toets + "_TOTAAL/printenscan/antwoorden_" + jaar + "_" +  toets +"_TOTAAL.qsf"
    qsfTargetFilename= "../" + jaar + "_" +  toets + "/antwoorden_" + jaar + "_" +  toets +".qsf"
    shutil.copyfile(qsfSourceFilename, qsfTargetFilename)
    
def genereerPuntenBestand(jaar,toets,sessie,onderdelen,regelFeedbackgroep,regelGeslaagd):
    #lees punten van TOTAAL
    outputFolder_onderdeel = "/output_" + jaar+ "_" +  toets + "_TOTAAL/"
    puntenFilename= "../" + jaar + "_" +  toets + "_TOTAAL" + outputFolder_onderdeel + "punten_" + jaar + "_" +  toets + "_TOTAAL.xls"
    punten_onderdeel = pd.read_excel(puntenFilename)#,dtype=str)

    columns_punten = [punten_onderdeel.columns[x] for x in [0,1,3,4,5]]
    punten_compose=punten_onderdeel[columns_punten]
    namen_nieuw = ["nummer","TOTAAL","juist","fout","blanco"]
    punten_compose.columns=namen_nieuw
    
    punten_compose.insert(1,"ijkingstoetssessie",numpy.ones(punten_compose.shape[0]) * sessie)
    punten_compose.insert(1,"ijkID",[""]* punten_compose.shape[0])
    punten_compose.insert(0,"Voornaam",[""]* punten_compose.shape[0])
    punten_compose.insert(0,"Naam",[""]* punten_compose.shape[0])
    
   
    for onderdeel in onderdelen:
        onderdeelFolder = "../" + jaar + "_" +  toets + "_"+ onderdeel
        if not os.path.exists(onderdeelFolder):
             print("Error: folder " + onderdeelFolder + " does not exist")
             sys.exit()
        toetsnaamOnderdeel = toets + "_" + onderdeel
        outputFolder_onderdeel = "/output_" + jaar + "_" + toetsnaamOnderdeel + "/"
        puntenFilename= onderdeelFolder + outputFolder_onderdeel + "punten_" + jaar + "_" +  toets + "_" + onderdeel + ".xls"
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
        
    
    geslaagdVariabele=bepaalGeslaagd(punten_compose,regelGeslaagd)
    punten_compose.insert(5,"Geslaagd",geslaagdVariabele)
    
    feedbackgroep=bepaalFeedbackGroep(punten_compose,regelFeedbackgroep)
    punten_compose.insert(5,"FeedbackGroep",feedbackgroep)
        
    #punten_compose["nummer","FeedbackGroep","Geslaagd"].to_excel("../" + jaar + "_" +  toets +"/resultaten.xlsx",sheet_name="punten",index=False)
    punten_compose.to_excel("../" + jaar + "_" +  toets +"/resultaten_"+ jaar + "_" + toets + ".xls",sheet_name="punten",index=False)

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

def bepaalGeslaagd(df,regelGeslaagd):
    geslaagdVariabele = [""]* df.shape[0]
    
    if regelGeslaagd == "geslaagdTotaal":            
        geslaagdGroep = (df["TOTAAL"].values>=10)
        nietGeslaagdGroep = [not x for x in geslaagdGroep]
    if regelGeslaagd == "ia":            
        geslaagdGroep = (df["TOTAAL"].values>=10) & (df["scoreB"].values>=10)
        nietGeslaagdGroep = [not x for x in geslaagdGroep]

    geslaagdVariabele = numpy.where(geslaagdGroep,True,geslaagdVariabele)
    geslaagdVariabele = numpy.where(nietGeslaagdGroep,False,geslaagdVariabele)
    
    #print(geslaagdVariabele)
    return geslaagdVariabele
    