# -*- coding: utf-8 -*-
"""
Created on 9/7/2021

@author: tdelaet

assumptions:
    -There is a folder with name jaar_toets, e.g. 2021_ia21
     The folder contains:
        - OMR folder with:
            at least 1 excel file with OMR output:
                jaar_toets_OMRoutput_instelling.xslx e.g. 2021_ia21_OMRoutput_Brussel.xslx
        - onderdelen folder with:
            - maxScores_jaar_toets.txt (e.g. maxScores 2021_ia21.txt)
                that file contains the maximum scores for the whole test and the subparts
                 maxScore_TOTAAL, maxScore_A, maxScore_B, maxScore_C
                 bvb: 20,10,20,10 
                 conditions: (have to be integers separated by commas) en (length = 1 + number of subparts	)
            - sleutel_jaar_toets.txt (e.g. sleutel_2021_ia21.txt)
                that file contains the correct answers for the total test
                A,B,B,D,A,C ... 
                conditions: (have to be letters separated by commas) en (length = number of questions in total test	)
            - for each subpart a file jaar_toets_subpart.txt
                that contains the questions that are in that subpart
                e.g. 19,20,21,22
            - IF different series/permutations are used: permutatie_jaar_toets.txt (e.g. permutatie_2021_ia21.txt)
    - Below fill in variables jaar, toets, sessie, editie, aantal_onderdelen, numSeries, numAlternatives, blankAnswer, verwerking
"""
from xlrd import open_workbook
import string
import numpy
import matplotlib.pyplot as plt
from xlwt import Workbook
import matplotlib
import os
import sys

import checkInputVariables
import supportFunctions
import writeResults
import leesSleutelEnPermutaties
import voorbereidingOnderdelen
import afwerkingOnderdelen


### Variables to fill in
jaar = "2022"
sessie = 22
toets = "bi" 
editie= "juli "+ jaar
aantal_onderdelen = 3
numSeries=2 # number of series TODO lezen van file

#dw
#regelFeedbackgroep = "dw"     #A als (TOTAAL >=9 )   ; B als (5<=TOTAAL<9 )    ; C als (TOTAAL<5 ) 
#regelGeslaagd =  "geslaagdTotaal" #A als (TOTAAL >=10)  

#ir
#regelFeedbackgroep =  "geslaagdTotaal" #A als (TOTAAL >=10)  
#regelGeslaagd =  "geslaagdTotaal" #A als (TOTAAL >=10) 


#ia
#regelFeedbackgroep = "ia"      #A als (TOTAAL >=10 & scoreB>=10)    
#regelGeslaagd = "ia"      #geslaagd als (TOTAAL >=10 & scoreB>=10)  

#ww
#regelFeedbackgroep =  "iedereenA"
#regelGeslaagd =  "geslaagdTotaal" #A als (TOTAAL >=10) 

#bi
regelFeedbackgroep =  "geslaagdTotaal" #A als (TOTAAL >=10)  
regelGeslaagd =  "geslaagdTotaal" #A als (TOTAAL >=10) 

#in/id/ib
#regelFeedbackgroep =  "geslaagdTotaal" #A als (TOTAAL >=10)  
#regelGeslaagd =  "geslaagdTotaal" #A als (TOTAAL >=10) 


numAlternatives = 4 #number of alternatives

#instellingen = ["Leuven","Kortrijk","Gent","Brussel","Howest"]
#instellingen = ["LEUVEN","LD","GENT","BRUSSEL","GK","Kulak"]
#instellingen = ["Leuven","Gent","Brussel","Kortrijk"]
#instellingen = ["Leuven","Gent","Brussel","Kortrijk","online"]
instellingen = ["all"]
#instellingen = ["all","online"]
#instellingen = ["Leuven"]

blankAnswer = "X" 

verwerking = "text" #als sleutel en permutatie als txt gegeven
#verwerking = "tex" #als sleutel en permutatie als tex zijn gegeven

toets = toets + str(sessie)
# do you want to write a feedback excel, one sheet per student?
writeFeedbackStudents = False

# code from here
if numSeries==1:
     permutationsUsed = False
else:
     permutationsUsed = True

onderdelen = voorbereidingOnderdelen.voorbereidingOnderdelen(jaar,toets,permutationsUsed,aantal_onderdelen,instellingen)

for onderdeel in (["TOTAAL"] + onderdelen):
    print("-------------------------------------------------")
    print("VERWERKING ONDERDEEL " + onderdeel)
    print("-------------------------------------------------")
    toetsnaamOnderdeel = toets + "_" + onderdeel
    folder_onderdeel = "../" + jaar + "_" +  toetsnaamOnderdeel
    # number of questions is length of the sleutel/correct answers
    numQuestions = len(numpy.loadtxt(folder_onderdeel + "/sleutel_"  + jaar +"_" + toetsnaamOnderdeel +".txt",delimiter=',',dtype="str",ndmin=1))
    print("aantal vragen "+ str(numQuestions))
    maxTotalScore = numpy.loadtxt(folder_onderdeel + "/maxScore_" + jaar+ "_"+ toetsnaamOnderdeel + ".txt",delimiter=',',dtype="int",ndmin=1)[0]
    print ("maximum score "+ str(maxTotalScore))
    
    #nameFile = "../OMR/test" #name of excel file with scanned forms
    nameSheet = "outputScan" #sheet name of excel file with scanned forms
    
    nameFile = folder_onderdeel + "/OMR/"+ jaar + "_" +  toetsnaamOnderdeel + "_OMRoutput" #name of excel file with scanned forms
    
    texinputFolder = folder_onderdeel + "/texinput/"        
    texoutputFolder = folder_onderdeel + "/texoutput/"
    if verwerking == "tex":
        if not os.path.exists(texoutputFolder):
            os.makedirs(texoutputFolder)    
    
    #where output of processing is saved
    outputFolder_onderdeel = folder_onderdeel + "/output_" + jaar + "_" + toetsnaamOnderdeel + "/"
    if not os.path.exists(outputFolder_onderdeel):
        os.makedirs(outputFolder_onderdeel)
    #where output for print en scan is save
    outputFolder_onderdeel_ps = folder_onderdeel + "/printenscan/"
    if not os.path.exists(outputFolder_onderdeel_ps):
        os.makedirs(outputFolder_onderdeel_ps)
    
    distributionList = [0.35,0.5,0.6,0.7,0.8,0.9]
    bordersDistributionStudentsLow =  [int(maxTotalScore*x) for x in distributionList]#for counting how many students get <=7,10 ...
    bordersDistributionStudentsHigh = bordersDistributionStudentsLow#for counting how many students get >=7,10 ...
    
    ############################
    #create list of expected content of scan file
    content = ["ijkID","vragenreeks"]
    
    for question in range(1,numQuestions+1):
            name = "Vraag" + str(question)
            content.append(name)
    ###########################

    ############################
    if verwerking == 'tex':
        #correct answers
        correctAnswers = leesSleutelEnPermutaties.leesSleutel(jaar,toetsnaamOnderdeel,texinputFolder)
        #permutations
        if numSeries == 1:
            permutations = numpy.zeros((1,numQuestions))
            for question in range(0,numQuestions):
                permutations[0,question] = question + 1
        else:
            permutations = leesSleutelEnPermutaties.leesPermutaties(jaar,toetsnaamOnderdeel,numSeries,texinputFolder)
    else:
        #correct answers
        correctAnswers = numpy.loadtxt(folder_onderdeel + "/sleutel_" + jaar+ "_"+ toetsnaamOnderdeel+ ".txt",delimiter=',',dtype="str",ndmin=1)
        #permutations
        if numSeries == 1:
            permutations = numpy.zeros((1,numQuestions))
            for question in range(0,numQuestions):
                permutations[0,question] = question + 1
        else:
            permutations = numpy.loadtxt(folder_onderdeel +  "/permutatie_" + jaar+ "_"+ toetsnaamOnderdeel+ ".txt",delimiter=',',dtype=numpy.float)
     
    print("sleutel: ")
    print(correctAnswers)
    print("permutaties: ")
    print(permutations)

    ### Important for tex processing, of texinputFolders does not exists, produces "empty" list
    #name of questions
    nameQuestions = leesSleutelEnPermutaties.leesNamenVragen(jaar,toetsnaamOnderdeel,texinputFolder,numQuestions)
    #classification of questions
    classificationQuestionsMod = leesSleutelEnPermutaties.leesClassificatieVragen(jaar,toetsnaamOnderdeel,texinputFolder,numQuestions)
    #categorie of questions
    categorieQuestions = leesSleutelEnPermutaties.leesCategorieVragen(jaar,toetsnaamOnderdeel,texinputFolder,numQuestions)

    
    #numpy.savetxt(outputFolder_onderdeel + "permutatie_"+ jaar +"_" + toetsnaamOnderdeel + ".txt",permutations,delimiter=',',fmt="%i")
    ############################
    ############################
    
    plt.close("all")

    #letters of answer alternatives
    alternatives = list(string.ascii_uppercase)[0:numAlternatives]
    
         
    if not( checkInputVariables.checkInputVariables(nameFile,nameSheet,numQuestions,numAlternatives,numSeries,correctAnswers,permutations,nameQuestions,instellingen,classificationQuestionsMod,categorieQuestions)):
        print ("ERROR found in input variables"   )
        sys.exit()
    
    deelnemers_all = []      
    scoreQuestionsAllPermutations_all = []
    numQuestionsAlternatives_all = []
    scoreQuestionsIndicatedSeries_all = []
    numberCorrect_all = []
    numberCorrect_wrong = []
    numberCorrect_blank = []
    totalScoreDifferentPermutations_all = []
    totalScore_all = []
    averageScore_all = []
    medianScore_all = []
    standardDeviation_all = []
    percentagePass_all = []
    columnSeries_all = []
    matrixAnswers_all = []
    numParticipants_all = []
    scoreCategories_all = []
    numberCorrectAnswers_all = []
    numberWrongAnswers_all = []
    numberBlankAnswers_all = []
      
    for instelling in instellingen:  
        counter = 0
        print ("instelling: " + instelling)
        # read file and get sheet
        book= open_workbook(nameFile+"_"+ instelling+".xlsx")
        sheet = book.sheet_by_name(nameSheet)
            
        #number of rows and columns
        num_rows = sheet.nrows;
        num_cols = sheet.ncols;
     
        #number of participants = number of rows-1
        numParticipants = num_rows-1;
        
        #Load the first row => name indicating 
        firstRow =  sheet.row(0) 
        firstRowValues = sheet.row_values(0)
        
        content_colNrs = supportFunctions.giveContentColNrs(content, sheet);
        
        #prepare a matrix to store the score of a student with all possible permutations
        scoreQuestionsIndicatedSeries= numpy.zeros((numParticipants,numQuestions))
        
        # prepare output excels
        outputbook = Workbook(style_compression=2)
        outputbookperm = Workbook(style_compression=2)
        outputStudentbook = Workbook(style_compression=2)
        outputResults = Workbook(style_compression=2)
        if writeFeedbackStudents:
            outputFeedbackbook = Workbook(style_compression=2)
        
        name = "ijkID"
        studentenNrCol= content_colNrs[content.index(name)]
        deelnemers=sheet.col_values(studentenNrCol,1,num_rows)
        
        if not supportFunctions.checkForUniqueParticipants(deelnemers):
            print ("ERROR: Duplicate participants found")
            sys.exit()
        
        name = "vragenreeks"
        #get the column in which the vragenreeks is stored
        colNrSerie = content_colNrs[content.index(name)]
        #get the series for the participants (so skip for row with name of first row)
        columnSeries=sheet.col_values(colNrSerie,1,num_rows)
                  
        # get matrix of answers
        matrixAnswers = supportFunctions.getMatrixAnswers(sheet,content,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs)  
        supportFunctions.checkMatrixAnswers(matrixAnswers,alternatives,blankAnswer)
        
        #get the score for all permutations for each of the questions
        scoreQuestionsAllPermutations= supportFunctions.calculateScoreAllPermutations(sheet,matrixAnswers,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs)     
        #scoreQuestionsAllPermutations= supportFunctions.calculateScoreAllPermutations_old(sheet,content,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs)     
        numQuestionsAlternatives = supportFunctions.getNumberAlternatives(sheet,content,permutations,columnSeries,scoreQuestionsIndicatedSeries,alternatives,blankAnswer,content_colNrs)
        
        #get the scores for the indicated series
        scoreQuestionsIndicatedSeries, averageScoreQuestions, numberCorrectAnswers, numberWrongAnswers, numberBlankAnswers =  supportFunctions.getScoreQuestionsIndicatedSeries(scoreQuestionsAllPermutations,columnSeries,numAlternatives)
        
        #get the overall statistics
        totalScore, averageScore, medianScore, standardDeviation, percentagePass = supportFunctions.getOverallStatistics(scoreQuestionsIndicatedSeries,maxTotalScore)
        
        #get all the scores for the different permutations
        totalScoreDifferentPermutations = supportFunctions.calculateTotalScoreDifferentPermutations(scoreQuestionsAllPermutations,maxTotalScore)
        #print totalScoreDifferentPermutations
        
        #get the average score for the different question categories
        scoreCategories = supportFunctions.getScoreCategories(scoreQuestionsIndicatedSeries,categorieQuestions)
        
        numParticipantsSeries, averageScoreSeries, medianScoreSeries, standardDeviationSeries, percentagePassSeries, averageScoreQuestionsDifferentSeries = supportFunctions.getOverallStatisticsDifferentSeries(totalScoreDifferentPermutations,scoreQuestionsIndicatedSeries,columnSeries,maxTotalScore)
        
        totalScoreUpper,totalScoreMiddle,totalScoreLower,averageScoreUpper, averageScoreMiddle, averageScoreLower, averageScoreQuestionsUpper, averageScoreQuestionsMiddle, averageScoreQuestionsLower,numQuestionsAlternativesUpper,numQuestionsAlternativesMiddle,numQuestionsAlternativesLower, scoreQuestionsUpper, scoreQuestionsMiddle, scoreQuestionsLower,numUpper, numMiddle, numLower= supportFunctions.calculateUpperLowerStatistics(matrixAnswers,content,columnSeries,totalScore,scoreQuestionsIndicatedSeries,correctAnswers,alternatives,blankAnswer,content_colNrs,permutations)
         
        distributionStudentsHigh,distributionStudentsLow = supportFunctions.getDistributionStudents(totalScore,bordersDistributionStudentsLow,bordersDistributionStudentsHigh)
        
        #only plot and write to files if there are multiple instellingen. Otherwise "geheel" is same as single instelling, producing double output.
        if (len(instellingen)!=1):
            ## WRITING THE OUTPUT TO A FILE
            writeResults.write_results(outputbook,outputbookperm,numQuestions,correctAnswers,alternatives,blankAnswer,
                              maxTotalScore,content,content_colNrs,
                              columnSeries,deelnemers,
                              numParticipants,
                              totalScore,percentagePass,
                              scoreQuestionsIndicatedSeries,
                              totalScoreDifferentPermutations,
                              medianScore,
                              standardDeviation,
                              averageScore,averageScoreUpper,averageScoreMiddle,averageScoreLower,
                              averageScoreQuestions,averageScoreQuestionsUpper,averageScoreQuestionsMiddle,averageScoreQuestionsLower,
                              averageScoreQuestionsDifferentSeries,
                              numUpper,numMiddle,numLower,
                              numParticipantsSeries,
                              averageScoreSeries,medianScoreSeries,standardDeviationSeries,percentagePassSeries,
                              numQuestionsAlternatives, numQuestionsAlternativesUpper, numQuestionsAlternativesMiddle, numQuestionsAlternativesLower,
                              nameQuestions,classificationQuestionsMod,categorieQuestions,
                              bordersDistributionStudentsLow,bordersDistributionStudentsHigh,
                              distributionStudentsLow,distributionStudentsHigh)
            
                         
            ## WRITING A FILE TO UPLOAD TO TOLEDO WITH THE GRADES
            writeResults.write_scoreStudents(outputStudentbook,"punten",permutations,numParticipants,deelnemers, numQuestions,numAlternatives,content,content_colNrs,totalScore,scoreQuestionsIndicatedSeries,columnSeries,matrixAnswers,numberCorrectAnswers, numberWrongAnswers, numberBlankAnswers)           
            #writeResults.write_resultsFile(outputResults,"resultaten",permutations,numParticipants,deelnemers, numQuestions,numAlternatives,content,content_colNrs,totalScore,scoreQuestionsIndicatedSeries,columnSeries,matrixAnswers,numberCorrectAnswers, numberWrongAnswers, numberBlankAnswers)                   

            writeResults.write_scoreCategoriesStudents(outputStudentbook,"percentageCategorien",deelnemers, totalScore, categorieQuestions, scoreCategories)
            
            outputFolder_instelling = outputFolder_onderdeel + instelling + "/"
            if not os.path.exists(outputFolder_instelling):
                os.makedirs(outputFolder_instelling)    
            outputbook.save(outputFolder_instelling + 'output_'  + jaar + "_" +  toets + "_" +instelling+'.xls') 
            outputbookperm.save(outputFolder_instelling + 'output_permutations_'  + jaar + "_" +  toets + "_" +instelling+'.xls') 
            outputStudentbook.save(outputFolder_instelling + 'punten_'  + jaar + "_" +  toets + "_" +instelling+'.xls') 
            #outputResults.save(outputFolder_instelling + 'resultaten_'  + jaar + "_" +  toets + "_" +instelling+'.xls') 
                              
            # plot the histogram of the total score
            plt.figure(figsize=(15, 5))
            n, bins, patches = plt.hist(totalScore,bins=numpy.arange(0-0.5,maxTotalScore+1,1))
            plt.title("histogram score " + instelling)
            plt.xlabel("score (max " + str(maxTotalScore)+ ")")
            plt.xlim([0-0.5,maxTotalScore+0.5])
            plt.xticks(numpy.arange(1,maxTotalScore+1))
            plt.ylabel("aantal studenten")
            plt.text(maxTotalScore, numpy.max(n)-2, 
                  'gemiddelde: ' + str(round(averageScore,2)) + "\n" +
                  'mediaan: ' + str(int(medianScore))  + "\n" +
                  'percentage geslaagd: ' + str(int(round(percentagePass,0))) + "%"  + "\n" +
                  'aantal deelnemers: ' + str(numParticipants)
                  ,
                horizontalalignment='right',
                verticalalignment='top',
                bbox=dict(facecolor='none', edgecolor='black', boxstyle='round,pad=1'))
            figManager = plt.get_current_fig_manager()
            #figManager.window.showMaximized()    
            plt.savefig(outputFolder_instelling + 'histogramGeheel_'  + jaar + "_" +  toets + "_" +instelling+'.png', bbox_inches='tight',dpi=300)
            
    
        deelnemers_all.append(deelnemers)
        scoreQuestionsAllPermutations_all.append(scoreQuestionsAllPermutations)
        numQuestionsAlternatives_all.append(numQuestionsAlternatives)
        scoreQuestionsIndicatedSeries_all.append(scoreQuestionsIndicatedSeries)
        totalScoreDifferentPermutations_all.append(totalScoreDifferentPermutations)
        totalScore_all.append(totalScore)
        averageScore_all.append(averageScore)
        medianScore_all.append(medianScore)
        standardDeviation_all.append(standardDeviation)
        percentagePass_all.append(percentagePass)
        columnSeries_all.append(columnSeries)
        matrixAnswers_all.append(matrixAnswers)
        numParticipants_all.append(numParticipants)
        scoreCategories_all.append(scoreCategories)
        numberCorrectAnswers_all.append(numberCorrectAnswers)
        numberWrongAnswers_all.append(numberWrongAnswers)
        numberBlankAnswers_all.append(numberBlankAnswers)

    
    deelnemers_tot = numpy.hstack(deelnemers_all)
    scoreQuestionsAllPermutations_tot = numpy.hstack(scoreQuestionsAllPermutations_all)
    numQuestionsAlternatives_tot = sum(numQuestionsAlternatives_all)
    scoreQuestionsIndicatedSeries_tot = numpy.vstack(scoreQuestionsIndicatedSeries_all)
    totalScoreDifferentPermutations_tot = numpy.vstack(totalScoreDifferentPermutations_all)
    columnSeries_tot = numpy.hstack(columnSeries_all)
    matrixAnswers_tot = numpy.vstack(matrixAnswers_all)
    numParticipants_stacked_tot = numpy.vstack(numParticipants_all)
    averageScore_stacked_tot = numpy.vstack(averageScore_all)
    medianScore_stacked_tot = numpy.vstack(medianScore_all)
    standardDeviation_stacked_tot = numpy.vstack(standardDeviation_all)
    percentagePass_stacked_tot  = numpy.vstack(percentagePass_all)
    numParticipants_tot = sum(numParticipants_stacked_tot)[0]
    scoreCategories_tot = numpy.hstack(scoreCategories_all)

    totalScore_tot, averageScore_tot, medianScore_tot, standardDeviation_tot, percentagePass_tot = supportFunctions.getOverallStatistics(scoreQuestionsIndicatedSeries_tot,maxTotalScore)
    numParticipantsSeries_tot, averageScoreSeries_tot, medianScoreSeries_tot, standardDeviationSeries_tot, percentagePassSeries_tot, averageScoreQuestionsDifferentSeries_tot = supportFunctions.getOverallStatisticsDifferentSeries(totalScoreDifferentPermutations_tot,scoreQuestionsIndicatedSeries_tot,columnSeries_tot,maxTotalScore)
    scoreQuestionsIndicatedSeries_tot, averageScoreQuestions_tot, numberCorrectAnswers_tot, numberWrongAnswers_tot, numberBlankAnswers_tot =  supportFunctions.getScoreQuestionsIndicatedSeries(scoreQuestionsAllPermutations_tot,columnSeries_tot,numAlternatives)
        
    totalScoreUpper_tot,totalScoreMiddle_tot,totalScoreLower_tot,averageScoreUpper_tot, averageScoreMiddle_tot, averageScoreLower_tot, averageScoreQuestionsUpper_tot, averageScoreQuestionsMiddle_tot, averageScoreQuestionsLower_tot,numQuestionsAlternativesUpper_tot,numQuestionsAlternativesMiddle_tot,numQuestionsAlternativesLower_tot, scoreQuestionsUpper_tot, scoreQuestionsMiddle_tot, scoreQuestionsLower_tot,numUpper_tot, numMiddle_tot, numLower_tot= supportFunctions.calculateUpperLowerStatistics(matrixAnswers_tot,content,columnSeries_tot,totalScore_tot,scoreQuestionsIndicatedSeries_tot,correctAnswers,alternatives,blankAnswer,content_colNrs,permutations)
    distributionStudentsHigh_tot,distributionStudentsLow_tot= supportFunctions.getDistributionStudents(totalScore_tot,bordersDistributionStudentsLow,bordersDistributionStudentsHigh)

    
    # write to excel_file
    outputbook = Workbook(style_compression=2)
    outputbookperm = Workbook(style_compression=2)
    outputStudentbook = Workbook(style_compression=2)  
    outputResults = Workbook(style_compression=2)  
    outputInstellingen = Workbook(style_compression=2)  
    if writeFeedbackStudents:
        outputFeedbackbook = Workbook(style_compression=2)
    outputFeedbackPlatformbook = Workbook(style_compression=2)
    outputDeelnemersLijst = Workbook(style_compression=2)
    

    ## WRITING THE OUTPUT TO A FILE
    writeResults.write_qsf(outputFolder_onderdeel_ps,numAlternatives,numQuestions,matrixAnswers_tot,correctAnswers,deelnemers_tot,columnSeries_tot,jaar,toetsnaamOnderdeel)
    writeResults.write_results(outputbook,outputbookperm,numQuestions,correctAnswers,alternatives,blankAnswer,
                      maxTotalScore,content,content_colNrs,
                      columnSeries_tot,deelnemers_tot,
                      numParticipants_tot,
                      totalScore_tot,percentagePass_tot,
                      scoreQuestionsIndicatedSeries_tot,
                      totalScoreDifferentPermutations_tot,
                      medianScore_tot,
                      standardDeviation_tot,
                      averageScore_tot,averageScoreUpper_tot,averageScoreMiddle_tot,averageScoreLower_tot,
                      averageScoreQuestions_tot,averageScoreQuestionsUpper_tot,averageScoreQuestionsMiddle_tot,averageScoreQuestionsLower_tot,
                      averageScoreQuestionsDifferentSeries_tot,
                      numUpper_tot,numMiddle_tot,numLower_tot,
                      numParticipantsSeries_tot,
                      averageScoreSeries_tot,medianScoreSeries_tot,standardDeviationSeries_tot,percentagePassSeries_tot,
                      numQuestionsAlternatives_tot, numQuestionsAlternativesUpper_tot, numQuestionsAlternativesMiddle_tot, numQuestionsAlternativesLower_tot,
                      nameQuestions,classificationQuestionsMod,categorieQuestions,
                      bordersDistributionStudentsLow,bordersDistributionStudentsHigh,distributionStudentsLow_tot,distributionStudentsHigh_tot
                      )    
    
    
    writeResults.write_scoreStudents(outputStudentbook,"punten",permutations,numParticipants_tot,deelnemers_tot, numQuestions,numAlternatives,content,content_colNrs,totalScore_tot,scoreQuestionsIndicatedSeries_tot,columnSeries_tot,matrixAnswers_tot,numberCorrectAnswers_tot,numberWrongAnswers_tot,numberBlankAnswers_tot)           
    writeResults.write_resultsFile(outputResults,"resultaten",permutations,numParticipants_tot,deelnemers_tot, numQuestions,numAlternatives,content,content_colNrs,totalScore_tot,scoreQuestionsIndicatedSeries_tot,columnSeries_tot,matrixAnswers_tot,numberCorrectAnswers_tot,numberWrongAnswers_tot,numberBlankAnswers_tot)           
    writeResults.write_overallStatisticsInstellingen(outputInstellingen,"instellingen",instellingen,numParticipants_tot,numParticipants_stacked_tot,averageScore_tot,averageScore_stacked_tot,medianScore_tot,medianScore_stacked_tot,standardDeviation_tot,standardDeviation_stacked_tot,percentagePass_tot,percentagePass_stacked_tot)
    #writeResults.write_scoreStudentsNonPermutated(outputStudentbook,"verwerking",numSeries,permutations,numParticipants,deelnemers, numQuestions,numAlternatives,alternatives,content,content_colNrs,totalScore,scoreQuestionsIndicatedSeries,columnSeries,matrixAnswers)
    writeResults.write_scoreStudentsNonPermutated(outputStudentbook,"punten_reeks1",permutations,numParticipants,deelnemers, numQuestions,numAlternatives,alternatives,content,content_colNrs,totalScore,scoreQuestionsIndicatedSeries,columnSeries,matrixAnswers)
    writeResults.write_scoreCategoriesStudents(outputStudentbook,"percentageCategorien",deelnemers_tot,totalScore_tot, categorieQuestions, scoreCategories_tot)
    

    if writeFeedbackStudents:
        writeResults.write_feedbackStudents(outputFeedbackbook,permutations,numParticipants_tot,deelnemers_tot, numQuestions,
                                        alternatives,numAlternatives,content,content_colNrs,
                                        totalScore_tot,scoreQuestionsIndicatedSeries_tot,columnSeries_tot,matrixAnswers_tot,
                                        categorieQuestions,scoreCategories_tot,
                                        averageScoreQuestions_tot,averageScoreQuestionsUpper_tot,averageScoreQuestionsMiddle_tot,averageScoreQuestionsLower_tot
                                        ,correctAnswers, numQuestionsAlternatives_tot)
    # writeResults.write_feedbackPlatform(outputFolder_onderdeel,permutations,numParticipants_tot,deelnemers_tot, numQuestions,
    #                                     alternatives,numAlternatives,content,content_colNrs,
    #                                     totalScore_tot,scoreQuestionsIndicatedSeries_tot,columnSeries_tot,matrixAnswers_tot,
    #                                     categorieQuestions,scoreCategories_tot,
    #                                     averageScoreQuestions_tot,averageScoreQuestionsUpper_tot,averageScoreQuestionsMiddle_tot,averageScoreQuestionsLower_tot
    #                                     ,correctAnswers, numQuestionsAlternatives_tot,blankAnswer) 
    writeResults.write_participantsList(outputDeelnemersLijst,"Beoordelingen",deelnemers_tot)

    outputbook.save(outputFolder_onderdeel + 'output' +'_controleerVoorKwaliteitToets_'  + jaar + "_" +  toetsnaamOnderdeel + '.xls')  
    outputbookperm.save(outputFolder_onderdeel + 'output_controleerVoorFouteReeksen_'  + jaar + "_" +  toetsnaamOnderdeel + '.xls')  
    if (len(instellingen)!=1):
        outputInstellingen.save(outputFolder_onderdeel + 'instellingen_'  + jaar + "_" +  toetsnaamOnderdeel + '.xls')  
    outputStudentbook.save(outputFolder_onderdeel + 'punten_'  + jaar + "_" +  toetsnaamOnderdeel + '.xls')  
    outputResults.save(outputFolder_onderdeel + '../resultaten_'  + jaar + "_" +  toetsnaamOnderdeel + '.xls')  
    outputDeelnemersLijst.save(outputFolder_onderdeel_ps + 'deelnemerslijst_KULoket_'  + jaar + "_" +  toetsnaamOnderdeel + '.xls')  
    if writeFeedbackStudents:
        outputFeedbackbook.save(outputFolder_onderdeel+ 'feedback_'  + jaar + "_" +  toetsnaamOnderdeel + '.xls')  

    def my_autopct(pct):
        total=sum(numParticipants_all)
        val=int(pct*total/100.0)
        return '{p:.2f}%  ({v:d})'.format(p=pct,v=val)
    
    ##################PLOTTING##################
    font = {'family' : 'normal',
            'size'   : 12}
    if (len(instellingen)!=1):
        # plot the pie diagram of the different locations
        plt.figure()
        labels = instellingen    
        plt.pie(numParticipants_all, labels=labels,
                        autopct=my_autopct, shadow=True, startangle=90)    
        plt.title('Aantal deelnemers', bbox={'facecolor':'0.8', 'pad':5})
        plt.savefig(outputFolder_onderdeel + 'verdelingDeelnemers_'  + jaar + "_" +  toetsnaamOnderdeel + '.png', bbox_inches='tight',dpi=300)
    
    # plot the histogram of the total score
    fig=plt.figure(figsize=(15, 5))
    ax=fig.add_subplot(111)
    n, bins, patches = plt.hist(totalScore_tot,bins=numpy.arange(0-0.5,maxTotalScore+1,1))
    plt.xlabel("score (max " + str(maxTotalScore)+ ")")
    plt.xlim([0-0.5,maxTotalScore+0.5])
    plt.xticks(numpy.arange(1,maxTotalScore+1))
    plt.ylabel("aantal studenten")       
    plt.text(0.966,0.9, 
              'gemiddelde: ' + str(round(averageScore_tot,2)) + "\n" +
              'mediaan: ' + str(int(medianScore_tot+1))  + "\n" +
              'percentage geslaagd: ' + str(int(round(percentagePass_tot,0))) + "%"  + "\n" +
              'aantal deelnemers: ' + str(numParticipants_tot)
              ,transform=ax.transAxes,
            horizontalalignment='right',
            verticalalignment='top',
            bbox=dict(facecolor='none', edgecolor='black', boxstyle='round,pad=1'),
            fontsize=12)     
    matplotlib.rc('font', **font)       
    #figManager = plt.get_current_fig_manager()
    #figManager.window.showMaximized()    
    plt.savefig(outputFolder_onderdeel + 'histogramGeheel_'  + jaar + "_" +  toetsnaamOnderdeel + '.png', bbox_inches='tight',dpi=300)
    if verwerking=="tex":
        plt.savefig(texoutputFolder + 'histogramGeheel_'  + jaar + "_" +  toetsnaamOnderdeel + '.png', bbox_inches='tight',dpi=300)
    
    # plot the histogram of the total score UML
    plt.figure(figsize=(15, 5))
    n, bins, patches = plt.hist([totalScoreUpper_tot,totalScoreMiddle_tot,totalScoreLower_tot],bins=numpy.arange(0-0.5,maxTotalScore+1,1), stacked=True, color=['g', 'b', 'r'])
    plt.title("histogram total score")
    plt.xlabel("score (max " + str(maxTotalScore)+ ")")
    plt.xlim([0-0.5,maxTotalScore+0.5])
    plt.ylabel("aantal studenten")
    plt.text(maxTotalScore, numpy.max(n)-0.5, 
              'gemiddelde: ' + str(round(averageScore_tot,2)) + "\n" +
              'mediaan: ' + str(int(medianScore_tot))  + "\n" +
              'percentage geslaagd: ' + str(int(round(percentagePass_tot,0))) + "%"  + "\n" +
              'aantal deelnemers: ' + str(numParticipants_tot) +"\n" +
              'Upper gemiddelde: ' + str(round(averageScoreUpper_tot,2))  + "\n" +
              'Middle gemiddelde: ' + str(round(averageScoreMiddle_tot,2)) + "\n" +
              'Lower gemiddelde: ' + str(round(averageScoreLower_tot,2)) 
              ,horizontalalignment='right'
              , verticalalignment='top'
              , bbox=dict(facecolor='none', edgecolor='black', boxstyle='round,pad=1'))
    figManager = plt.get_current_fig_manager()
    #figManager.window.showMaximized()            
    plt.savefig(outputFolder_onderdeel + 'histogramGeheelUML_'  + jaar + "_" +  toetsnaamOnderdeel + '.png', bbox_inches='tight',dpi=300)

    #plot histogram for different questions
    numColsPict = int(numpy.ceil(numpy.sqrt(numQuestions)))
    numRowsPict = int(numpy.ceil(numQuestions/numColsPict))
    if (numRowsPict*numColsPict < numQuestions):
        numRowsPict+=1
    fig, axes = plt.subplots(nrows=numRowsPict, ncols=numColsPict,figsize=(15, 15))
    fig.tight_layout() # Or equivalently,  "plt.tight_layout()"
    binsHist = numpy.array([-3.0/(2*(numAlternatives-1)),-1.0/(2*(numAlternatives-1)),0.5,1.5])
    for question in range(1,numQuestions+1):
        ax = plt.subplot(numRowsPict,numColsPict,question)
        n, bins, patches = plt.hist(scoreQuestionsIndicatedSeries_tot[:,question-1],bins=binsHist)
        plt.xticks([round(-1/(numAlternatives-1),2), 0,1])
        plt.title("vraag " + str(question))
        plt.xlabel("score")
        plt.xlim([-2.0/(numAlternatives-1),1+1.0/(numAlternatives-1)])
        plt.ylabel("aantal studenten")
    matplotlib.rc('font', **font)
    figManager = plt.get_current_fig_manager()
    #figManager.window.showMaximized()    
    plt.savefig(outputFolder_onderdeel + 'histogramVragen_'  + jaar + "_" +  toetsnaamOnderdeel + '.png', bbox_inches='tight',dpi=300)
       
    #plot histogram for different questions UML
    numColsPict = int(numpy.ceil(numpy.sqrt(numQuestions)))
    numRowsPict = int(numpy.ceil(numQuestions/numColsPict))
    if (numRowsPict*numColsPict < numQuestions):
        numRowsPict+=1
    fig, axes = plt.subplots(nrows=numRowsPict, ncols=numColsPict,figsize=(15, 15))
    fig.tight_layout() # Or equivalently,  "plt.tight_layout()"
    
    for question in range(1,numQuestions+1):
        ax = plt.subplot(numRowsPict,numColsPict,question)
        correctUpper =  sum(scoreQuestionsUpper_tot[:,question-1] == 1.0)/len(scoreQuestionsUpper_tot[:,question-1])
        correctMiddle = sum(scoreQuestionsMiddle_tot[:,question-1] == 1.0)/len(scoreQuestionsMiddle_tot[:,question-1])
        correctLower = sum(scoreQuestionsLower_tot[:,question-1] == 1.0)/len(scoreQuestionsLower_tot[:,question-1])                        
        plt.bar(["lower","middle","upper"],[correctLower,correctMiddle,correctUpper],width=0.8)
        plt.title("vraag " + str(question))
        plt.ylabel("%correct")
        plt.ylim([0,1])
    figManager = plt.get_current_fig_manager()
    #figManager.window.showMaximized()    
    plt.savefig(outputFolder_onderdeel + 'histogramVragenUML_'  + jaar + "_" +  toetsnaamOnderdeel + '.png', bbox_inches='tight',dpi=300)
    
    # #feedback file schrijven
    # fin = open(texinputFolder + 'feedbackdraft.tex','r')
    # fout= open(texoutputFolder + 'feedback.tex','w')
    # inhoud=fin.read()
    # inhoud=inhoud.replace('<editie>',editie)
    # inhoud=inhoud.replace('<aantal>',str(numParticipants_tot))
    # inhoud=inhoud.replace('<G>', str(int(distributionStudentsHigh_tot[1])))
    # inhoud=inhoud.replace('<N1>', str(round(distributionStudentsHigh_tot[5]/numParticipants_tot*100,1)))
    # inhoud=inhoud.replace('<N2>', str(round(distributionStudentsHigh_tot[4]/numParticipants_tot*100,1)))
    # inhoud=inhoud.replace('<N3>', str(round(distributionStudentsHigh_tot[3]/numParticipants_tot*100,1)))
    # inhoud=inhoud.replace('<N4>', str(round(distributionStudentsHigh_tot[2]/numParticipants_tot*100,1)))
    # inhoud=inhoud.replace('<N5>', str(round(distributionStudentsHigh_tot[1]/numParticipants_tot*100,1)))
    # inhoud=inhoud.replace('<N6>', str(round(distributionStudentsLow_tot[0]/numParticipants_tot*100,1)))
    # fout.write(inhoud)
    # fin.close()
    # fout.close()
    
    # #statistische gegevens in tex-file schrijven
    # nameFile = [[] for i in range(int(numQuestions))]
    # frapport = open(texoutputFolder + 'rapportinput.tex','w')
    # for vraag in range(0,numQuestions):
    #     percCorrectr = int(round(numQuestionsAlternatives_tot[vraag,alternatives.index(correctAnswers[vraag])]/numParticipants_tot*100,0))
    #     percBlankr = int(round(numQuestionsAlternatives_tot[vraag,numAlternatives]/numParticipants_tot*100,0))
    #     percUpperr = int(round(numQuestionsAlternativesUpper_tot[vraag,alternatives.index(correctAnswers[vraag])]/numUpper_tot*100,0))
    #     percLowerr = int(round(numQuestionsAlternativesLower_tot[vraag,alternatives.index(correctAnswers[vraag])]/numLower_tot*100,0))
    #     if not os.path.isfile(texinputFolder + nameQuestions[vraag] + '.tex'):
    #         fin = open(texinputFolder + 'vraagdraft.tex','r')
    #         nameFile[vraag]="vraag" + str(int(vraag+1))
    #     else:
    #         fin = open(texinputFolder + nameQuestions[vraag] + '.tex','r')
    #         nameFile[vraag]=nameQuestions[vraag]
    #     inhoud=fin.read()
    #     inhoud=inhoud.replace('<vraagnr>',str(int(vraag+1)))
    #     inhoud=inhoud.replace('<editie>',editie)
    #     inhoud=inhoud.replace('<aantal>',str(numParticipants_tot))
    #     inhoud=inhoud.replace('<juist>',str(percCorrectr))
    #     inhoud=inhoud.replace('<blanco>',str(percBlankr))
    #     ULABCD = 'upper/lower:'+str(percUpperr)+'/'+str(percLowerr)+'\\newline percentages ABCD:'
    #     ULABCD = ULABCD+ str(int(round(numQuestionsAlternatives_tot[vraag,0]/numParticipants_tot*100,0)))+'/'
    #     ULABCD = ULABCD+ str(int(round(numQuestionsAlternatives_tot[vraag,1]/numParticipants_tot*100,0)))+'/'
    #     ULABCD = ULABCD+ str(int(round(numQuestionsAlternatives_tot[vraag,2]/numParticipants_tot*100,0)))+'/'
    #     ULABCD = ULABCD+ str(int(round(numQuestionsAlternatives_tot[vraag,3]/numParticipants_tot*100,0)))
    #     inhoud=inhoud.replace('<ul>',ULABCD)
    #     fout= open(texoutputFolder + nameFile[vraag] + '_stat.tex','w')
    #     fout.write(inhoud)
    #     fin.close()
    #     fout.close()
    #     frapport.write("\\input{vraag" + str(int(vraag+1))  + "_stat}\n" )
    # frapport.close()
    
afwerkingOnderdelen.genereerPuntenBestand(jaar,toets,sessie,onderdelen,regelFeedbackgroep,regelGeslaagd)
afwerkingOnderdelen.kopieerQSF(jaar,toets)