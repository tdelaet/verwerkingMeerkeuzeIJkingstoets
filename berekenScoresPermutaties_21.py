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

import checkInputVariables
import supportFunctions
import writeResults
import leesSleutelEnPermutaties

jaar = "2021"
toets = "hw21_E"
editie= "juli "+ jaar
#verwerking = "tex" #als sleutel en permutatie als tex gegeven
verwerking = "text" #als sleutel en permutatie als txt gegeven

numQuestions = 3 # number of questions
maxTotalScore = 10 #maximum total score
numAlternatives = 4 #number of alternatives
numSeries= 1 # number of series
blankAnswer = "X"

#nameFile = "../OMR/test" #name of excel file with scanned forms
nameSheet = "outputScan" #sheet name of excel file with scanned forms


#instellingen = ["Leuven","Kortrijk","Gent","Brussel","Howest"]
#instellingen = ["LEUVEN","LD","GENT","BRUSSEL","GK","Kulak"]
#instellingen = ["Leuven","Gent","Brussel","Kortrijk"]
instellingen = ["all"]

writeFeedbackStudents = False

nameFile = "../" + jaar + "_" +  toets + "/OMR/"+ jaar + "_" +  toets + "_OMRoutput" #name of excel file with scanned forms


texinputFolder = "../" + jaar + "_" +  toets + "/texinput/"

outputFolder = "../" + jaar + "_" +  toets + "/output/"
if not os.path.exists(outputFolder):
    os.makedirs(outputFolder)
    
texoutputFolder = "../" + jaar + "_" +  toets + "/texoutput/"
if not os.path.exists(texoutputFolder):
    os.makedirs(texoutputFolder)    

bordersDistributionStudentsLow = [7,10,12,14,16,18] #for counting how many students get <=7,10 ...
bordersDistributionStudentsHigh = [7,10,12,14,16,18]#for counting how many students get >=7,10 ...

#######################
#write qsf-file
########################
fqsf= open(outputFolder + 'antwoorden.qsf','w')
fqsf.write('Snapshot,Participant,Vendor,Group')
for question in range(1,numQuestions+1):
    fqsf.write(',Q'+str(question))
fqsf.write('\n')

############################
#create list of expected content of scan file
content = ["ijkID","vragenreeks"]

for question in range(1,numQuestions+1):
        name = "Vraag" + str(question)
        content.append(name)
###########################

############################
############################
if verwerking == 'tex':
    #correct answers
    correctAnswers = leesSleutelEnPermutaties.leesSleutel(jaar,toets,texinputFolder)
    #permutations
    if numSeries == 1:
        permutations = numpy.zeros((1,numQuestions))
        for question in range(0,numQuestions):
            permutations[0,question] = question + 1
    else:
        permutations = leesSleutelEnPermutaties.leesPermutaties(jaar,toets,numSeries,texinputFolder)
else:
    correctAnswers = numpy.loadtxt("../" + jaar + "_" +  toets + "/sleutel_" + jaar+ "_"+ toets+ ".txt",delimiter=',',dtype="str",ndmin=1)
    #permutations
    if numSeries == 1:
        permutations = numpy.zeros((1,numQuestions))
        for question in range(0,numQuestions):
            permutations[0,question] = question + 1
    else:
        permutations = numpy.loadtxt("../" + jaar + "_" +  toets +  "/permutatie_" + jaar+ "_"+ toets+ ".txt",delimiter=',',dtype=numpy.float)
 
print("sleutel: ")
print(correctAnswers)
print("permutations: ")
print(permutations)

#name of questions
nameQuestions = leesSleutelEnPermutaties.leesNamenVragen(jaar,toets,texinputFolder,numQuestions)
#classification of questions
classificationQuestionsMod = leesSleutelEnPermutaties.leesClassificatieVragen(jaar,toets,texinputFolder,numQuestions)
#categorie of questions
categorieQuestions = leesSleutelEnPermutaties.leesCategorieVragen(jaar,toets,texinputFolder,numQuestions)


numpy.savetxt(outputFolder + "permutatie_"+ jaar +"_" +toets + ".txt",permutations,delimiter=',',fmt="%i")
############################
############################

plt.close("all")

#letters of answer alternatives
alternatives = list(string.ascii_uppercase)[0:numAlternatives]

     
if not( checkInputVariables.checkInputVariables(nameFile,nameSheet,numQuestions,numAlternatives,numSeries,correctAnswers,permutations,nameQuestions,instellingen,classificationQuestionsMod,categorieQuestions)):
    print ("ERROR found in input variables"   )

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
    print ("INSTELLING: " + instelling)
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
    
    #print columnSeries
    scoreQuestionsIndicatedSeries= numpy.zeros((numParticipants,numQuestions))
    
    # write to excel_file
    outputbook = Workbook(style_compression=2)
    outputbookperm = Workbook(style_compression=2)
    outputStudentbook = Workbook(style_compression=2)
    outputFeedbackbook = Workbook(style_compression=2)
    
    name = "ijkID"
    studentenNrCol= content_colNrs[content.index(name)]
    deelnemers=sheet.col_values(studentenNrCol,1,num_rows)
    
    if not supportFunctions.checkForUniqueParticipants(deelnemers):
        print ("ERROR: Duplicate participants found")
    
    name = "vragenreeks"
    #get the column in which the vragenreeks is stored
    colNrSerie = content_colNrs[content.index(name)]
    #get the series for the participants (so skip for row with name of first row)
    columnSeries=sheet.col_values(colNrSerie,1,num_rows)
    #write data to qsf
    for participant in range(0,numParticipants):
        fqsf.write(str(int(columnSeries[participant]))
        +',\"' + str(int(deelnemers[participant]))
        +'\",\"Gravic, Inc.\",\"auto\"')   
        antwoorden=sheet.row_values(1+participant,2,numQuestions+2)
        for vraag in range(0,numQuestions):
            fqsf.write(',' + str(antwoorden[vraag]))    
        fqsf.write('\n')
    # get matrix of answers
    matrixAnswers = supportFunctions.getMatrixAnswers(sheet,content,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs)  
    print(matrixAnswers)
    supportFunctions.checkMatrixAnswers(matrixAnswers,alternatives,blankAnswer)
    
    #get the score for all permutations for each of the questions
    scoreQuestionsAllPermutations= supportFunctions.calculateScoreAllPermutations(sheet,content,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs)     
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
    writeResults.write_scoreCategoriesStudents(outputStudentbook,"percentageCategorien",deelnemers, totalScore, categorieQuestions, scoreCategories)
    
    outputbook.save(outputFolder + 'output' +'_'+instelling+'.xls') 
    outputbookperm.save(outputFolder + 'output_permutations' +'_'+instelling+'.xls') 

    outputStudentbook.save(outputFolder + 'punten' +'_'+instelling+'.xls')
                      
    # plot the histogram of the total score
    plt.figure()
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
    plt.savefig(outputFolder + 'histogramGeheel'+ instelling + '.png', bbox_inches='tight',dpi=300)
    

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
outputInstellingen = Workbook(style_compression=2)  
if writeFeedbackStudents:
    outputFeedbackbook = Workbook(style_compression=2)
outputFeedbackPlatformbook = Workbook(style_compression=2)
outputDeelnemersLijst = Workbook(style_compression=2)

## WRITING THE OUTPUT TO A FILE
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
#writeResults.write_scoreStudentsNonPermutated(outputStudentbook,"verwerking",numSeries,permutations,numParticipants,deelnemers, numQuestions,numAlternatives,alternatives,content,content_colNrs,totalScore,scoreQuestionsIndicatedSeries,columnSeries,matrixAnswers)
writeResults.write_scoreStudentsNonPermutated(outputStudentbook,"punten_reeks1",permutations,numParticipants,deelnemers, numQuestions,numAlternatives,alternatives,content,content_colNrs,totalScore,scoreQuestionsIndicatedSeries,columnSeries,matrixAnswers)
writeResults.write_scoreCategoriesStudents(outputStudentbook,"percentageCategorien",deelnemers_tot,totalScore_tot, categorieQuestions, scoreCategories_tot)
writeResults.write_overallStatisticsInstellingen(outputInstellingen,"instellingen",instellingen,numParticipants_tot,numParticipants_stacked_tot,averageScore_tot,averageScore_stacked_tot,medianScore_tot,medianScore_stacked_tot,standardDeviation_tot,standardDeviation_stacked_tot,percentagePass_tot,percentagePass_stacked_tot)

if writeFeedbackStudents:
    writeResults.write_feedbackStudents(outputFeedbackbook,permutations,numParticipants_tot,deelnemers_tot, numQuestions,
                                    alternatives,numAlternatives,content,content_colNrs,
                                    totalScore_tot,scoreQuestionsIndicatedSeries_tot,columnSeries_tot,matrixAnswers_tot,
                                    categorieQuestions,scoreCategories_tot,
                                    averageScoreQuestions_tot,averageScoreQuestionsUpper_tot,averageScoreQuestionsMiddle_tot,averageScoreQuestionsLower_tot
                                    ,correctAnswers, numQuestionsAlternatives_tot)
writeResults.write_feedbackPlatform(outputFolder,permutations,numParticipants_tot,deelnemers_tot, numQuestions,
                                    alternatives,numAlternatives,content,content_colNrs,
                                    totalScore_tot,scoreQuestionsIndicatedSeries_tot,columnSeries_tot,matrixAnswers_tot,
                                    categorieQuestions,scoreCategories_tot,
                                    averageScoreQuestions_tot,averageScoreQuestionsUpper_tot,averageScoreQuestionsMiddle_tot,averageScoreQuestionsLower_tot
                                    ,correctAnswers, numQuestionsAlternatives_tot,blankAnswer) 
writeResults.write_participantsList(outputDeelnemersLijst,"Beoordelingen",deelnemers_tot)

outputbook.save(outputFolder + 'output' +'_geheel.xls')
outputbookperm.save(outputFolder + 'output_permutations' +'_geheel.xls') 
outputStudentbook.save(outputFolder + 'punten_geheel.xls') 
outputInstellingen.save(outputFolder + 'instellingen.xls')  
outputDeelnemersLijst.save(outputFolder + 'deelnemerslijst_KULoket.xls')  
if writeFeedbackStudents:
    outputFeedbackbook.save(outputFolder+ 'feedback'+'.xls')

def my_autopct(pct):
    total=sum(numParticipants_all)
    val=int(pct*total/100.0)
    return '{p:.2f}%  ({v:d})'.format(p=pct,v=val)
    
# plot the pie diagram of the different locations
plt.figure()
# The slices will be ordered and plotted counter-clockwise.
labels = instellingen

plt.pie(numParticipants_all, labels=labels,
                autopct=my_autopct, shadow=True, startangle=90)
                # The default startangle is 0, which would start
                # the Frogs slice on the x-axis.  With startangle=90,
                # everything is rotated counter-clockwise by 90 degrees,
                # so the plotting starts on the positive y-axis.

#plt.title('Aantal deelnemers', bbox={'facecolor':'0.8', 'pad':5})
plt.title('Aantal deelnemers', bbox={'facecolor':'0.8', 'pad':5})
plt.savefig(outputFolder + 'verdelingDeelnemers.png', bbox_inches='tight',dpi=300)


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
font = {'family' : 'normal',
        'size'   : 14}
matplotlib.rc('font', **font)       
#figManager = plt.get_current_fig_manager()
#figManager.window.showMaximized()    
plt.savefig(outputFolder + 'histogramGeheel.png', bbox_inches='tight',dpi=300)
plt.savefig(texoutputFolder + 'histogramGeheel.png', bbox_inches='tight',dpi=300)


# plot the histogram of the total score UML
plt.figure()
n, bins, patches = plt.hist([totalScoreUpper_tot,totalScoreMiddle_tot,totalScoreLower_tot],bins=numpy.arange(0-0.5,maxTotalScore+1,1), stacked=True, color=['g', 'b', 'r'])
plt.title("histogram total score")
plt.xlabel("score (max " + str(maxTotalScore)+ ")")
plt.xlim([0-0.5,maxTotalScore+0.5])
plt.ylabel("aantal studenten")


plt.text(maxTotalScore, numpy.max(n)-0.5, 
         'gemiddelde: ' + str(round(averageScore_tot,2)) + "\n" +
         'mediaan: ' + str(int(medianScore_tot))  + "\n" +
         'percentage geslaagd: ' + str(int(round(percentagePass_tot,0))) + "%"  + "\n" +
         'aantal deelnemers: ' + str(numParticipants_tot)
         ,
        horizontalalignment='right',
        verticalalignment='top',
        bbox=dict(facecolor='none', edgecolor='black', boxstyle='round,pad=1'))

plt.text(maxTotalScore, numpy.max(n)-7.5, 
         'Upper gemiddelde: ' + str(round(averageScoreUpper_tot,2)),
        horizontalalignment='right',
        verticalalignment='top')
#        bbox=dict(facecolor='none', edgecolor='green', boxstyle='round,pad=1'))
        
        
plt.text(maxTotalScore, numpy.max(n)-9.5, 
         'Middle gemiddelde: ' + str(round(averageScoreMiddle_tot,2)),
        horizontalalignment='right',
        verticalalignment='top')
        #bbox=dict(facecolor='none', edgecolor='blue', boxstyle='round,pad=1'))
        
plt.text(maxTotalScore, numpy.max(n)-11.5, 
         'Lower gemiddelde: ' + str(round(averageScoreLower_tot,2)),
        horizontalalignment='right',
        verticalalignment='top')
#        bbox=dict(facecolor='none', edgecolor='red', boxstyle='round,pad=1'))
figManager = plt.get_current_fig_manager()
#figManager.window.showMaximized()            
plt.savefig(outputFolder + 'histogramGeheelUML.png', bbox_inches='tight',dpi=300)


#plot histogram for different questions
numColsPict = int(numpy.ceil(numpy.sqrt(numQuestions)))
#print numColsPict
numRowsPict = int(numpy.ceil(numQuestions/numColsPict)) +1
#print numRowsPict
fig, axes = plt.subplots(nrows=numRowsPict, ncols=numColsPict)
fig.tight_layout() # Or equivalently,  "plt.tight_layout()"

binsHist = numpy.array([-3.0/(2*(numAlternatives-1)),-1.0/(2*(numAlternatives-1)),0.5,1.5])

for question in range(1,numQuestions+1):
    ax = plt.subplot(numRowsPict,numColsPict,question)
    n, bins, patches = plt.hist(scoreQuestionsIndicatedSeries_tot[:,question-1],bins=binsHist)
    plt.xticks([-1/(numAlternatives-1), 0,1])
    plt.title("vraag " + str(question))
    plt.xlabel("score")
    plt.xlim([-2.0/(numAlternatives-1),1+1.0/(numAlternatives-1)])
    plt.ylabel("aantal studenten")


font = {'family' : 'normal',
        'size'   : 12}

matplotlib.rc('font', **font)
figManager = plt.get_current_fig_manager()
#figManager.window.showMaximized()    
plt.savefig(outputFolder + 'histogramVragen.png', bbox_inches='tight',dpi=300)

#plot histogram for different questions
numColsPict = int(numpy.ceil(numpy.sqrt(numQuestions)))
#print numColsPict
numRowsPict = int(numpy.ceil(numQuestions/numColsPict)) +1
#print numRowsPict
fig, axes = plt.subplots(nrows=numRowsPict, ncols=numColsPict)
fig.tight_layout() # Or equivalently,  "plt.tight_layout()"

for question in range(1,numQuestions+1):
    ax = plt.subplot(numRowsPict,numColsPict,question)
    n, bins, patches = plt.hist([scoreQuestionsUpper_tot[:,question-1], scoreQuestionsMiddle_tot[:,question-1], scoreQuestionsLower_tot[:,question-1]],bins=binsHist, stacked=True,  label=['Upper', 'Middle', 'Lower'],color=['g','b','r'])
    plt.xticks([-1/(numAlternatives-1), 0,1])
    plt.title("vraag " + str(question))
    plt.xlabel("score")
    plt.xlim([-2.0/(numAlternatives-1),1+1.0/(numAlternatives-1)])
    plt.ylabel("aantal studenten")
    plt.legend(loc=2,prop={'size':6})
figManager = plt.get_current_fig_manager()
#figManager.window.showMaximized()    
plt.savefig(outputFolder + 'histogramVragenUML.png', bbox_inches='tight',dpi=300)





#qsf-file afsluiten
fqsf.write('1,\"999999\",\"Gravic, Inc.\",\"auto\"')   
for vraag in range(0,numQuestions):
    if correctAnswers[vraag]=='A': 
        sleutel='1'
    if correctAnswers[vraag]=='B': 
        sleutel='2'
    if correctAnswers[vraag]=='C': 
        sleutel='3'
    if correctAnswers[vraag]=='D': 
        sleutel='4'
    if correctAnswers[vraag]=='E': 
        sleutel='5'
    fqsf.write(',' + sleutel)   
fqsf.write('\n')
fqsf.close()


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

fpunten =  open(outputFolder + toets +'_punten_upload','w')
fpunten.write("\"Alle studenten\", , , , , \n")
fpunten.write("\"naam\",\"voornaam\",\"nummer\",\"ijkID\",\"Datum Examen\",\"TOTAAL\" \n")
for participant in range(0,numParticipants_tot):
    fpunten.write("-,-," + str(int(deelnemers_tot[participant])) + ",-,-,"+ str(int(totalScore_tot[participant])) + "\n") 
fpunten.close()

