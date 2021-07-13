import numpy
import os
import pandas as pd
import sys

def voorbereidingOnderdelen(jaar,toets,permutationsUsed,aantal_onderdelen,instellingen):
    #prepare main outputfolder
    outputFolder = "../" + jaar + "_" +  toets
    if not os.path.exists(outputFolder):
        os.makedirs(outputFolder)

    #construct list of letters for subparts (a,b,c,...)
    onderdelen=[]
    for letter in range(97,97+aantal_onderdelen):
        onderdelen.append(chr(letter).capitalize())

    outputFolderTotaal = "../" + jaar + "_" +  toets + "_TOTAAL"
    if not os.path.exists(outputFolderTotaal):
        os.makedirs(outputFolderTotaal)
        
    #prepare outputfolders for subparts
    for onderdeel in onderdelen:
        outputFolderOnderdeel = outputFolder + "_"+ onderdeel
        if not os.path.exists(outputFolderOnderdeel):
            os.makedirs(outputFolderOnderdeel)
        outputFolderOnderdeelPrintEnScan = outputFolder + "_"+ onderdeel + "/printenscan"
        if not os.path.exists(outputFolderOnderdeelPrintEnScan):
            os.makedirs(outputFolderOnderdeelPrintEnScan)
        #save wich questions are in subpart
        vragen_onderdeel = numpy.loadtxt(outputFolder + "/onderdelen/" + jaar+ "_"+ toets+ "_"+ onderdeel + ".txt",delimiter=',',dtype="int",ndmin=1)
        numpy.savetxt(outputFolderOnderdeel + "/vragen_" + jaar+ "_"+ toets+ "_"+ onderdeel + ".txt",[vragen_onderdeel],delimiter=',',fmt="%i")

    # prepare correct answers for subparts    
    sleutelOnderdelen(jaar,toets,onderdelen,outputFolder,outputFolderTotaal)
    # prepare permutations for subparts
    if permutationsUsed:
        permutatieOnderdelen(jaar,toets,onderdelen,outputFolder,outputFolderTotaal)
    # prepare maxScore of subparts
    maxScoreOnderdelen(jaar,toets,onderdelen,outputFolder,outputFolderTotaal)
    # prepare OMR for subports    
    OMROnderdelen(jaar,toets,onderdelen,instellingen,outputFolder,outputFolderTotaal)
    return onderdelen

def sleutelOnderdelen(jaar,toets,onderdelen,outputFolder,outputFolderTotaal):
    #get sleutel and save to TOTAAL folder
    correctAnswers = numpy.loadtxt(outputFolder + "/onderdelen/sleutel_" + jaar+ "_"+ toets+ ".txt",delimiter=',',dtype="str")
    #save sleutel in TOTAAL
    numpy.savetxt(outputFolderTotaal + "/sleutel_" + jaar+ "_"+ toets + "_TOTAAL.txt",[correctAnswers],delimiter=',',fmt="%s")
    
    #prepare sleutel for subparts
    for onderdeel in onderdelen:
        vragen_onderdeel = numpy.loadtxt("../" + jaar + "_" +  toets + "/onderdelen/" + jaar+ "_"+ toets+ "_"+ onderdeel.capitalize() + ".txt",delimiter=',',dtype="int",ndmin=1)

        # get sleutel with just questions of onderdeel
        if (vragen_onderdeel.ndim==0):
            correctAnswers_loc = correctAnswers[vragen_onderdeel]
        else:
            correctAnswers_loc = [correctAnswers[x-1] for x in vragen_onderdeel]
                
        # save sleutel of questions of subpart
        outputFolderOnderdeel = outputFolder + "_"+ onderdeel
        numpy.savetxt(outputFolderOnderdeel + "/sleutel_" + jaar+ "_"+ toets+ "_"+ onderdeel.capitalize() + ".txt",[correctAnswers_loc],delimiter=',',fmt="%s")

def maxScoreOnderdelen(jaar,toets,onderdelen,outputFolder,outputFolderTotaal):
    #get sleutel and save to main folder
    maxScores = numpy.loadtxt(outputFolder + "/onderdelen/maxScores_" + jaar+ "_"+ toets+ ".txt",delimiter=',',dtype="int",ndmin=1)
    if not onderdelen: #geen onderdelen
        if ( len(maxScores) != 1 ):
            print ("ERROR: het bestand "+  "/onderdelen/maxScores_" + jaar+ "_"+ toets+ ".txt" + " bevat niet het juiste aantal maximum scores. Het moet er " + len(onderdelen + 1) + " bevatten")
            sys.exit()
    else:
        if ( len(maxScores) != ( len(onderdelen) +1) ):
            print ("ERROR: het bestand "+  "/onderdelen/maxScores_" + jaar+ "_"+ toets+ ".txt" + " bevat niet het juiste aantal maximum scores. Het moet er " + len(onderdelen + 1) + " bevatten")
            sys.exit()
    #numpy.savetxt(outputFolder + "/maxScores_" + jaar+ "_"+ toets + ".txt",[maxScores],delimiter=',',fmt="%s")
    
    counter = 0

    #save maxScore TOTAAL
    numpy.savetxt(outputFolderTotaal + "/maxScore_" + jaar+ "_"+ toets+ "_TOTAAL.txt",[maxScores[counter]],delimiter=',',fmt="%s")
    counter = counter +1

    #save maxScore for subparts
    for onderdeel in onderdelen:
        maxScore_onderdeel = [maxScores[counter]]
        counter = counter +1
        outputFolderOnderdeel = outputFolder + "_"+ onderdeel
        numpy.savetxt(outputFolderOnderdeel + "/maxScore_" + jaar+ "_"+ toets+ "_"+ onderdeel.capitalize() + ".txt",maxScore_onderdeel,delimiter=',',fmt="%s")


def permutatieOnderdelen(jaar,toets,onderdelen, outputFolder,outputFolderTotaal):
    permutations = numpy.loadtxt("../" + jaar + "_" +  toets + "/onderdelen/permutatie_" + jaar+ "_"+ toets+ ".txt",delimiter=',',dtype="str")
    numpy.savetxt(outputFolderTotaal + "/permutatie_" + jaar+ "_"+ toets + "_TOTAAL.txt",permutations,delimiter=',',fmt="%s")
    for onderdeel in onderdelen:
       
        vragen_onderdeel = numpy.loadtxt("../" + jaar + "_" +  toets + "/onderdelen/" + jaar+ "_"+ toets+ "_"+ onderdeel.capitalize() + ".txt",delimiter=',',dtype="int",ndmin=1)
      
        # get permutatie with just questions of onderdeel
        if (vragen_onderdeel.ndim==0):
            permutations_loc = permutations[vragen_onderdeel]
        else:
            permutations_loc = [permutations[:,x-1] for x in vragen_onderdeel]
        
        outputFolderOnderdeel = outputFolder + "_"+ onderdeel
        outputFolderOnderdeelPrintEnScan = outputFolder + "_"+ onderdeel +"/printenscan"
        #subtract lowest number such that starts with question1
        #TODO: redo will only work if subsequent numbers in subparts
        permutations_loc = list(map(list, zip(*permutations_loc)))
        permutations_loc2 = [ [int(y)-int(min(permutations_loc[0]))+1 for y in x] for x in permutations_loc]
        numpy.savetxt(outputFolderOnderdeel + "/permutatie_" + jaar+ "_"+ toets+ "_"+ onderdeel.capitalize() + ".txt",permutations_loc2,delimiter=',',fmt="%s")
        numpy.savetxt(outputFolderOnderdeelPrintEnScan + "/permutatie_" + jaar+ "_"+ toets+ "_"+ onderdeel.capitalize() + ".txt",permutations_loc2,delimiter=',',fmt="%s")

        
# Deelt OMR op in onderdelen en schrijf die weg in map OMR in map van elk onderdeel
def OMROnderdelen(jaar,toets,onderdelen,instellingen,outputFolder,outputFolderTotaal):
    for instelling in instellingen: 
        OMRfilename = outputFolder + "/OMR/" + jaar+ "_"+ toets+ "_OMRoutput_" + instelling + ".xlsx"
        if not os.path.exists(OMRfilename):
            print ("ERROR: het bestand "+  OMRfilename + " bestaat niet")
            sys.exit()
        OMR = pd.read_excel(OMRfilename,dtype=str)
        OMR["vragenreeks"] = OMR["vragenreeks"].astype(str).astype(int)

        
        outputFolderOMR = outputFolder + "_TOTAAL/OMR"
        outputOMR= outputFolderOMR + "/" + jaar+ "_"+ toets+  "_TOTAAL_OMRoutput_" + instelling + ".xlsx"

        if not os.path.exists(outputFolderOMR):
            os.makedirs(outputFolderOMR)
        OMR.to_excel(outputOMR,sheet_name="outputScan",index=False)
        
        for onderdeel in onderdelen:
            #prepare folder for OMR
            outputFolderOMR = outputFolder + "_"+ onderdeel.capitalize() + "/OMR"
            if not os.path.exists(outputFolderOMR):
                os.makedirs(outputFolderOMR)
           
            vragen_onderdeel = numpy.loadtxt(outputFolder + "/onderdelen/" + jaar+ "_"+ toets+ "_"+ onderdeel.capitalize() + ".txt",delimiter=',',dtype="int",ndmin=1)
          
            #get OMR with just questions of onderdeel
            namen_onderdeel = ["ijkID","vragenreeks"]
        
            counter=1
            namen_onderdeelNieuw= ["ijkID","vragenreeks"]
            for x in vragen_onderdeel:
                namen_onderdeel.append("Vraag"+str(x))
                namen_onderdeelNieuw.append("Vraag"+str(counter))
                counter=counter+1
            
            outputOMR= outputFolderOMR + "/" + jaar+ "_"+ toets+  "_"+ onderdeel.capitalize() +  "_OMRoutput_" + instelling + ".xlsx"
            OMR_onderdeel = OMR[namen_onderdeel]
            #rename questions to Vraag1,Vraag2, ...
            OMR_onderdeel.columns=namen_onderdeelNieuw
            
            OMR_onderdeel.to_excel(outputOMR,sheet_name="outputScan",index=False)
