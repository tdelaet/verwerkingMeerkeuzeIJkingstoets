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

import checkInputVariables
import supportFunctions
import writeResults
import leesSleutelEnPermutaties

#nameFile = "../OMR/OMRoutput.xlsx" #name of excel file with scanned forms
#nameSheet = "outputScan" #sheet name of excel file with scanned forms

nameFile = "../OMR/Sept2013.xlsx" #name of excel file with scanned forms
nameSheet = "IR" #sheet name of excel file with scanned forms
numQuestions = 35 # number of questions
numAlternatives = 5 #number of alternatives
maxTotalScore = 20 #maximum total score
numSeries=4 # number of series
blankAnswer = "X"

jaar = "2013"
toets = "ir2"


#correct answers
correctAnswers = leesSleutelEnPermutaties.leesSleutel(jaar,toets)
#permutations
permutations = leesSleutelEnPermutaties.leesPermutaties(jaar,toets,numSeries)

#permutations = numpy.loadtxt("../permutatie_ir2.txt",delimiter=',',dtype=numpy.int32)

plt.close("all")

#letters of answer alternatives
alternatives = list(string.ascii_uppercase)[0:numAlternatives]
#numbers of answer alternatives
#alternatives = numpy.arange(1,numAlternatives+1)

############################
#create list of expected content of scan file
content = ["ijkID","vragenreeks"]
for question in xrange(1,numQuestions+1):
        name = "Vraag" + str(question)
        content.append(name)
#print content
###########################
        
if not( checkInputVariables.checkInputVariables(nameFile,nameSheet,numQuestions,numAlternatives,numSeries,correctAnswers,permutations)):
     print "ERROR found in input variables"   

        
  
# read file and get sheet
book= open_workbook(nameFile)
sheet = book.sheet_by_name(nameSheet)
    
#number of rows and columns
num_rows = sheet.nrows;
num_cols = sheet.ncols;

#number of participants = number of rows-1
numParticipants = num_rows-1;

#Load the first row => name indicating 
firstRow =  sheet.row(0) 
firstRowValues = sheet.row_values(0)
#print firstRowValues

content_colNrs = supportFunctions.giveContentColNrs(content, sheet);

            
#print columnSeries
scoreQuestionsIndicatedSeries= numpy.zeros((numParticipants,numQuestions))


# write to excel_file
outputbook = Workbook(style_compression=2)
outputStudentbook = Workbook(style_compression=2)


name = "ijkID"
studentenNrCol= content_colNrs[content.index(name)]
deelnemers=sheet.col_values(studentenNrCol,1,num_rows)

if not supportFunctions.checkForUniqueParticipants(deelnemers):
    print "ERROR: Duplicate participants found"


name = "vragenreeks"
#get the column in which the vragenreeks is stored
colNrSerie = content_colNrs[content.index(name)]
#get the series for the participants (so skip for row with name of first row)
columnSeries=sheet.col_values(colNrSerie,1,num_rows)


#get the score for all permutations for each of the questions
scoreQuestionsAllPermutations= supportFunctions.calculateScoreAllPermutations(sheet,content,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs)     
numQuestionsAlternatives = supportFunctions.getNumberAlternatives(sheet,content,permutations,columnSeries,scoreQuestionsIndicatedSeries,alternatives,blankAnswer,content_colNrs)


matrixAnswers = supportFunctions.getMatrixAnswers(sheet,content,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs)     


#get the scores for the indicated series
scoreQuestionsIndicatedSeries, averageScoreQuestions =  supportFunctions.getScoreQuestionsIndicatedSeries(scoreQuestionsAllPermutations,columnSeries)

#get the overall statistics
totalScore, averageScore, medianScore, standardDeviation, percentagePass = supportFunctions.getOverallStatistics(scoreQuestionsIndicatedSeries,maxTotalScore)
#print totalScore
#print averageScore
#print medianScore
#print standardDeviation
#print percentagePass

totalScoreDifferentPermutations = supportFunctions.calculateTotalScoreDifferentPermutations(scoreQuestionsAllPermutations,maxTotalScore)
#print totalScoreDifferentPermutations

numParticipantsSeries, averageScoreSeries, medianScoreSeries, standardDeviationSeries, percentagePassSeries, averageScoreQuestionsDifferentSeries = supportFunctions.getOverallStatisticsDifferentSeries(totalScoreDifferentPermutations,scoreQuestionsIndicatedSeries,columnSeries,maxTotalScore)

totalScoreUpper,totalScoreMiddle,totalScoreLower,averageScoreUpper, averageScoreMiddle, averageScoreLower, averageScoreQuestionsUpper, averageScoreQuestionsMiddle, averageScoreQuestionsLower,numQuestionsAlternativesUpper,numQuestionsAlternativesMiddle,numQuestionsAlternativesLower, scoreQuestionsUpper, scoreQuestionsMiddle, scoreQuestionsLower,numUpper, numMiddle, numLower= supportFunctions.calculateUpperLowerStatistics(sheet,content,columnSeries,totalScore,scoreQuestionsIndicatedSeries,correctAnswers,alternatives,blankAnswer,content_colNrs,permutations)
 


## WRITING THE OUTPUT TO A FILE
writeResults.write_results(outputbook,numQuestions,correctAnswers,alternatives,blankAnswer,
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
                  numQuestionsAlternatives, numQuestionsAlternativesUpper, numQuestionsAlternativesMiddle, numQuestionsAlternativesLower
                  )
                  
## WRITING A FILE TO UPLOAD TO TOLEDO WITH THE GRADES
writeResults.write_scoreStudents(outputStudentbook,"punten",permutations,numParticipants,deelnemers, numQuestions,numAlternatives,content,content_colNrs,totalScore,scoreQuestionsIndicatedSeries,columnSeries,matrixAnswers)           
                  

# plot the histogram of the total score
plt.figure()
n, bins, patches = plt.hist(totalScore,bins=numpy.arange(0-0.5,maxTotalScore+1,1))
plt.title("histogram score")
plt.xlabel("score (max " + str(maxTotalScore)+ ")")
plt.xlim([0-0.5,maxTotalScore+0.5])
plt.xticks(numpy.arange(1,21))
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

plt.savefig('histogramGeheel.png', bbox_inches='tight')

# plot the histogram of the total score UML
plt.figure()
n, bins, patches = plt.hist([totalScoreUpper,totalScoreMiddle,totalScoreLower],bins=numpy.arange(0-0.5,maxTotalScore+1,1), stacked=True, color=['g', 'b', 'r'])
plt.title("histogram total score")
plt.xlabel("score (max " + str(maxTotalScore)+ ")")
plt.xlim([0-0.5,maxTotalScore+0.5])
plt.ylabel("aantal studenten")


plt.text(maxTotalScore, numpy.max(n)-0.5, 
         'gemiddelde: ' + str(round(averageScore,2)) + "\n" +
         'mediaan: ' + str(int(medianScore))  + "\n" +
         'percentage geslaagd: ' + str(int(round(percentagePass,0))) + "%"  + "\n" +
         'aantal deelnemers: ' + str(numParticipants)
         ,
        horizontalalignment='right',
        verticalalignment='top',
        bbox=dict(facecolor='none', edgecolor='black', boxstyle='round,pad=1'))

plt.text(maxTotalScore, numpy.max(n)-7.5, 
         'Upper gemiddelde: ' + str(round(averageScoreUpper,2)),
        horizontalalignment='right',
        verticalalignment='top')
#        bbox=dict(facecolor='none', edgecolor='green', boxstyle='round,pad=1'))
        
        
plt.text(maxTotalScore, numpy.max(n)-9.5, 
         'Middle gemiddelde: ' + str(round(averageScoreMiddle,2)),
        horizontalalignment='right',
        verticalalignment='top')
        #bbox=dict(facecolor='none', edgecolor='blue', boxstyle='round,pad=1'))
        
plt.text(maxTotalScore, numpy.max(n)-11.5, 
         'Lower gemiddelde: ' + str(round(averageScoreLower,2)),
        horizontalalignment='right',
        verticalalignment='top')
#        bbox=dict(facecolor='none', edgecolor='red', boxstyle='round,pad=1'))
        
plt.savefig('histogramGeheelUML.png', bbox_inches='tight')


#plot histogram for different questions
numColsPict = int(numpy.ceil(numpy.sqrt(numQuestions)))
#print numColsPict
numRowsPict = int(numpy.ceil(numQuestions/numColsPict)) +1
#print numRowsPict
fig, axes = plt.subplots(nrows=numRowsPict, ncols=numColsPict)
fig.tight_layout() # Or equivalently,  "plt.tight_layout()"

binsHist = numpy.array([-3.0/(2*(numAlternatives-1)),-1.0/(2*(numAlternatives-1)),0.5,1.5])

for question in xrange(1,numQuestions+1):
    ax = plt.subplot(numRowsPict,numColsPict,question)
    n, bins, patches = plt.hist(scoreQuestionsIndicatedSeries[:,question-1],bins=binsHist)
    plt.xticks([-1/(numAlternatives-1), 0,1])
    plt.title("vraag " + str(question))
    plt.xlabel("score")
    plt.xlim([-2.0/(numAlternatives-1),1+1.0/(numAlternatives-1)])
    plt.ylabel("aantal studenten")
figManager = plt.get_current_fig_manager()
figManager.window.showMaximized()    
plt.savefig('histogramVragen.png', bbox_inches='tight')

#plot histogram for different questions
numColsPict = int(numpy.ceil(numpy.sqrt(numQuestions)))
#print numColsPict
numRowsPict = int(numpy.ceil(numQuestions/numColsPict)) +1
#print numRowsPict
fig, axes = plt.subplots(nrows=numRowsPict, ncols=numColsPict)
fig.tight_layout() # Or equivalently,  "plt.tight_layout()"

for question in xrange(1,numQuestions+1):
    ax = plt.subplot(numRowsPict,numColsPict,question)
    n, bins, patches = plt.hist([scoreQuestionsUpper[:,question-1], scoreQuestionsMiddle[:,question-1], scoreQuestionsLower[:,question-1]],bins=binsHist, stacked=True,  label=['Upper', 'Middle', 'Lower'],color=['g','b','r'])
    plt.xticks([-1/(numAlternatives-1), 0,1])
    plt.title("vraag " + str(question))
    plt.xlabel("score")
    plt.xlim([-2.0/(numAlternatives-1),1+1.0/(numAlternatives-1)])
    plt.ylabel("aantal studenten")
    plt.legend(loc=2,prop={'size':6})
figManager = plt.get_current_fig_manager()
figManager.window.showMaximized()    
plt.savefig('histogramVragenUML.png', bbox_inches='tight')



outputbook.save('output.xls') 
outputStudentbook.save('punten.xls')