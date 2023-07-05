import os
import pandas as pd
import sys
import shutil
import numpy

        
def kopieerQSF(jaar,toets,outputFolder):
    qsfSourceFilename= outputFolder + "_TOTAAL/printenscan/antwoorden_" + jaar + "_" +  toets +"_TOTAAL.qsf"
    qsfTargetFilename= outputFolder + "/antwoorden_" + jaar + "_" +  toets +".qsf"
    shutil.copyfile(qsfSourceFilename, qsfTargetFilename)
    
def genereerZIPs(jaar,toets,sessie,onderdelen,outputFolder):
    onderdelen_loc=onderdelen.copy()
    onderdelen_loc.insert(0,"TOTAAL")
    #print(outputFolder)
    for onderdeel in onderdelen_loc:
        #onderdeelFolder = os.path.dirname(os.getcwd()) + "\\" + jaar + "_" +  toets + "_"+ onderdeel + "\\"
        onderdeelFolder = outputFolder + "_"+ onderdeel
        toetsnaamOnderdeel = toets + "_" + onderdeel
        outputFolder_onderdeelFull =  onderdeelFolder + "/output_" + jaar + "_" + toetsnaamOnderdeel + "/"
        #print("Onderdeel " + onderdeel + "   folder: " + outputFolder_onderdeelFull)
        if not os.path.exists(outputFolder_onderdeelFull):
             print("Error: folder " + outputFolder_onderdeelFull + " does not exist")
             sys.exit()
        zipToCreate = outputFolder+ "_"+ onderdeel
        #print(zipToCreate)
        #print(outputFolder_onderdeelFull)
        shutil.make_archive(zipToCreate,"zip",outputFolder_onderdeelFull)
    
def genereerPuntenBestand(jaar,toets,sessie,onderdelen,regelFeedbackgroep,regelGeslaagd,maxScores,outputFolder):
    #print("begin genereerPUntenBestand")
    #lees punten van TOTAAL
    outputFolder_onderdeel = outputFolder +  "_TOTAAL/" + "/output_" + jaar+ "_" +  toets + "_TOTAAL/"
    #print("genereerPuntenbestand " + outputFolder_onderdeel )
    puntenFilename= outputFolder_onderdeel + "punten_" + jaar + "_" +  toets + "_TOTAAL.xls"
    punten_onderdeel = pd.read_excel(puntenFilename)#,dtype=str)

    columns_punten = [punten_onderdeel.columns[x] for x in [0,1,3,4,5]]
    punten_compose=punten_onderdeel[columns_punten]
    namen_nieuw = ["nummer","TOTAAL","juist","fout","blanco"]
    punten_compose.columns=namen_nieuw
    
    
    punten_compose.insert(1,"ijkingstoetssessie",(numpy.ones(punten_compose.shape[0]) * sessie).astype(int))
    punten_compose.insert(1,"ijkID",[""]* punten_compose.shape[0])
    punten_compose.insert(0,"Voornaam",[""]* punten_compose.shape[0])
    punten_compose.insert(0,"Naam",[""]* punten_compose.shape[0])
    #print(punten_compose.dtypes)
    
    #print("test")

    for onderdeel in onderdelen:
        #onderdeelFolder = "../" + jaar + "_" +  toets + "_"+ onderdeel
        onderdeelFolder = outputFolder + "_"+ onderdeel
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
    #print("test")

    for letter in range(97+len(onderdelen),97+8):
        legeOnderdelen.append(chr(letter).capitalize())
    for onderdeel in legeOnderdelen:
        namen_nieuw = ["score" + onderdeel,"juist"+ onderdeel,"fout"+ onderdeel,"blanco"+ onderdeel]
        df = pd.DataFrame(columns=namen_nieuw)
        punten_compose[namen_nieuw] = df[namen_nieuw]
        
    
    geslaagdVariabele=bepaalGeslaagd(punten_compose,regelGeslaagd,maxScores)
    punten_compose.insert(5,"Geslaagd",geslaagdVariabele)
    
    feedbackgroep=bepaalFeedbackGroep(punten_compose,regelFeedbackgroep,maxScores)
    punten_compose.insert(5,"FeedbackGroep",feedbackgroep)
    
    #print("test")
    #punten_compose["nummer","FeedbackGroep","Geslaagd"].to_excel("../" + jaar + "_" +  toets +"/resultaten.xlsx",sheet_name="punten",index=False)
    #print(punten_compose)
    punten_compose.to_csv(outputFolder +"/resultaten_"+ jaar + "_" + toets + ".csv", index = False) 
    #print("tussen")
    punten_compose.to_excel(outputFolder +"/resultaten_"+ jaar + "_" + toets + ".xls",sheet_name="punten",index=False)
    #print("end genereerPUntenBestand")
    
def bepaalFeedbackGroep(df,regelFeedbackgroep,maxScores):
    #print("feedbackgroup begin")
    feedbackgroep = [""]* df.shape[0]
    feedbackgroepA = [False]* df.shape[0]
    feedbackgroepB = [False]* df.shape[0]
    feedbackgroepC = [False]* df.shape[0]
    feedbackgroepD = [False]* df.shape[0]
    feedbackgroepE = [False]* df.shape[0]
    feedbackgroepF = [False]* df.shape[0]
    
    #iedereen feedbackgroep A
    if regelFeedbackgroep == "iedereenA":
        feedbackgroepA = [True]* df.shape[0]
    #feedbackgroepA als geslaagd op totaal, anders feedbackgroepB
    if regelFeedbackgroep == "geslaagdTotaal":            
        feedbackgroepA = (df["TOTAAL"].values>=maxScores[0]/2)
        feedbackgroepB = [not x for x in feedbackgroepA]
    #feedbackgroepA als totaal geslaagd en score B geslaagd, anders feedbackgroepB
    if regelFeedbackgroep == "ia":            
        feedbackgroepA = (df["TOTAAL"].values>=maxScores[0]/2) & (df["scoreB"].values>=maxScores[2]/2)
        feedbackgroepB = [not x for x in feedbackgroepA]
    #feedbackgroepA als score >=10; feedbackgroepB als score tussen 5 en 10; feedbackgroepC als score <5
    if regelFeedbackgroep == "dw":            
        feedbackgroepA = (df["TOTAAL"].values>=10)
        feedbackgroepB = (df["TOTAAL"].values<10) & (df["TOTAAL"].values>=5)
        feedbackgroepC = (df["TOTAAL"].values<5)
    #feedbackgroepA als score >=10; feedbackgroepB als 4<= score_TOTAAL < 10, feedbackgroepC als score score_TOTAAL<=3 
    if regelFeedbackgroep == "bi":            
        feedbackgroepA = (df["TOTAAL"].values>=12)
        feedbackgroepB = (df["TOTAAL"].values<12) & (df["TOTAAL"].values>=10)
        feedbackgroepC = (df["TOTAAL"].values<10)        
        #feedbackgroep A score_TOTAAL >=12; 
        #feedbackgroep B 10 <= score_TOTAAL<12;
        #feedbackgroep C score_TOTAAL<10
    if regelFeedbackgroep == "bwfa":            
        feedbackgroepA = (df["TOTAAL"].values>=10)
        feedbackgroepB = (df["TOTAAL"].values<10) & (df["TOTAAL"].values>6)
        feedbackgroepC = (df["TOTAAL"].values<=6)        
        #feedbackgroep A score_TOTAAL >=10 
        #feedbackgroep B score_TOTAAL <10 AND score TOTAAL > 6
        #feedbackgroep C score_TOTAAL <=6
    if regelFeedbackgroep == "ib":
        feedbackgroepA = (df["TOTAAL"].values>=12)
        feedbackgroepB = (df["TOTAAL"].values<12) & (df["TOTAAL"].values>=10)
        feedbackgroepC = (df["TOTAAL"].values<10) & (df["TOTAAL"].values>5)
        feedbackgroepD = (df["TOTAAL"].values<=5) 
        #feedbackgroep A score_Totaal>=12;
        #feedbackgroep B 10<=score_Totaal<12;
        #feedbackgroep C 5<score_Totaal<10;
        #feedbackgroep D score_Totaal<=5
    
    feedbackgroep = numpy.where(feedbackgroepA,"A",feedbackgroep)
    feedbackgroep = numpy.where(feedbackgroepB,"B",feedbackgroep)
    feedbackgroep = numpy.where(feedbackgroepC,"C",feedbackgroep)
    feedbackgroep = numpy.where(feedbackgroepD,"D",feedbackgroep)
    feedbackgroep = numpy.where(feedbackgroepE,"E",feedbackgroep)
    feedbackgroep = numpy.where(feedbackgroepF,"F",feedbackgroep)
    #print("feedbackgroep end")
    return feedbackgroep

def bepaalGeslaagd(df,regelGeslaagd,maxScores):
    #print("geslaagd begin")
    geslaagdVariabele = [""]* df.shape[0]
    
    if regelGeslaagd == "geslaagdTotaal":            
        geslaagdGroep = (df["TOTAAL"].values>=maxScores[0]/2)
        nietGeslaagdGroep = [not x for x in geslaagdGroep]
    if regelGeslaagd == "ia":            
        geslaagdGroep = (df["TOTAAL"].values>=maxScores[0]/2) & (df["scoreB"].values>=maxScores[2]/2)
        nietGeslaagdGroep = [not x for x in geslaagdGroep]
    #A als (score_TOTAAL>=10 AND score_wiskunde>=8)
    if regelGeslaagd == "wf":            
        geslaagdGroep = (df["TOTAAL"].values>=10) & (df["scoreB"].values>=8)
        nietGeslaagdGroep = [not x for x in geslaagdGroep]

    geslaagdVariabele = numpy.where(geslaagdGroep,True,geslaagdVariabele)
    geslaagdVariabele = numpy.where(nietGeslaagdGroep,False,geslaagdVariabele)
    
    #print("geslaagd end")
    return geslaagdVariabele
    