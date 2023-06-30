# -*- coding: utf-8 -*-
"""
Created on Wed May 21 14:54:46 2014

@author: tdelaet
"""

import numpy
import sys

def round2(x):
    """Numpy rounds x.5 to nearest even integer. To emulate SAS/SPSS, 
    which both round x.5 *up* to nearest integer, use this function.
    """
#    y = x - numpy.floor(x)
#    for i in numpy.arange(0,len(y)):
#        if (0 < y[i] < 0.5):
#            x[i] = numpy.floor(x[i])
#        else:
#            x[i] = numpy.ceil(x[i])
    x = numpy.floor(numpy.round(x,6)+0.5)    
    return x

def common_elements(list1, list2):
    return list(set(list1) & set(list2))
        
#def all_indices(value, qlist):
#    indices = []
#    idx = -1
#    while True:
#        try:
#            idx = qlist.index(value, idx+1)
#            indices.append(idx)
#        except ValueError:
#            break
#    return indices

def giveContentColNrs(content_loc, sheet_loc):
    content_colNrs_loc = [0] * len(content_loc);
    firstRowValues_loc = sheet_loc.row_values(0)
    indexC_loc =0
    # check if all content is present
    for x in content_loc:
        try:
            content_colNrs_loc[indexC_loc] = firstRowValues_loc.index(x)
            indexC_loc += 1
        except ValueError:
            print ("the expected content " + x + " is not present in the selected sheet")
            sys.exit()
            break;
    return content_colNrs_loc
    
def checkForUniqueParticipants(particpants):
    setd = set([x for x in particpants if particpants.count(x) > 1])
    if setd:
        print ("Duplicate participants found: " + str(setd))
        sys.exit()
        return False
    else:
        return True
    
def getMatrixAnswers(sheet_loc,contentBook_loc,correctAnswers_loc,permutations_loc,alternatives_loc,numParticipants_loc,columnSeries_loc,content_colNrs_loc):
    # Get the matrix of answers of the students
    numQuestions_loc = len(correctAnswers_loc)
    answers_loc= numpy.array(range(numParticipants_loc*numQuestions_loc),dtype=str).reshape(numParticipants_loc,numQuestions_loc)
    numAlternatives_loc = len(alternatives_loc)
   
    counterColumn = 0
    for question_loc in range(1,numQuestions_loc+1):
        name_question_serie1 = "Vraag" + str(question_loc)
        colNr_loc = content_colNrs_loc[contentBook_loc.index(name_question_serie1)]
        columnQuestion_loc=sheet_loc.col_values(colNr_loc,1,numParticipants_loc+1)

        answers_loc[:,counterColumn] = columnQuestion_loc;
        counterColumn+=1
    # replace OMR output 1, 2, 3 , 4 , 5 , 6 with A, B, C, D, E, X
    for alternative in range(0,numAlternatives_loc):
        letter = chr(97+alternative).capitalize()
        answers_loc = numpy.where(answers_loc==str(alternative+1), letter, answers_loc)
    print( answers_loc)
    #the blank answers who are thus equal to numAlternatives + 1 => replace by X
    #letter="X"
    #answers_loc = numpy.where(answers_loc==str(numAlternatives_loc+1), letter, answers_loc)
    return answers_loc

def calculateScoreAllPermutations(sheet_loc,matrixAnswers_loc,correctAnswers_loc,permutations_loc,alternatives_loc,numParticipants_loc,columnSeries_loc,content_colNrs_loc):
    # Calculate the score for each permutation and for each question 
    numSeries_loc = len(permutations_loc)
    numQuestions_loc = len(correctAnswers_loc)
    numAlternatives_loc = len(alternatives_loc)
    scoreQuestionsAllPermutations_loc= numpy.zeros((numSeries_loc,numParticipants_loc,numQuestions_loc))
   
    #Calculate score for all permutations
                 
    for question_loc in range(1,numQuestions_loc+1):
        #print "----------------------"
        #print "question " + str(question_loc)
        columnQuestion_loc=matrixAnswers_loc[:,question_loc-1]
        counter_alternative = 0;
        for permutation in range(1,numSeries_loc+1):
            numQuestionPermutations_loc = int(permutations_loc[permutation-1][question_loc-1])
            correctAnswer = correctAnswers_loc[numQuestionPermutations_loc-1]
            wrongAnswers = [x for x in alternatives_loc if x != correctAnswer]
            indicesCorrectAnswer_loc = [x for x in range(numParticipants_loc) if columnQuestion_loc[x]==correctAnswer]
            indicesWrongAnswer_loc = [x for x in range(numParticipants_loc) if (columnQuestion_loc[x] in set(wrongAnswers))]
            #correctAnswers +1
            scoreQuestionsAllPermutations_loc[permutation-1,indicesCorrectAnswer_loc,numQuestionPermutations_loc-1]+=1.0
            #wrong answers -1/(numAlternatives-1)
            scoreQuestionsAllPermutations_loc[permutation-1,indicesWrongAnswer_loc,numQuestionPermutations_loc-1]-= 1.0/(float(numAlternatives_loc)-1.0)  
            #blank answers => do nothing
        counter_alternative+=1
    return scoreQuestionsAllPermutations_loc

def calculateScoreAllPermutations_old(sheet_loc,contentBook_loc,correctAnswers_loc,permutations_loc,alternatives_loc,numParticipants_loc,columnSeries_loc,content_colNrs_loc):
    # Calculate the score for each permutation and for each question 
    numSeries_loc = len(permutations_loc)
    numQuestions_loc = len(correctAnswers_loc)
    numAlternatives_loc = len(alternatives_loc)
    scoreQuestionsAllPermutations_loc= numpy.zeros((numSeries_loc,numParticipants_loc,numQuestions_loc))
   
   #Calculate score for all permutations
                 
    for question_loc in range(1,numQuestions_loc+1):
        #print "----------------------"
        #print "question " + str(question_loc)
        counter_alternative = 0;

        name_question_serie1 = "Vraag" + str(question_loc)
        colNr_loc = content_colNrs_loc[contentBook_loc.index(name_question_serie1)]
        columnQuestion_loc=sheet_loc.col_values(colNr_loc,1,numParticipants_loc+1)
        #TODO: replace with matrixAnswers
        columnQuestion_loc = ['A' if x=='1'  else x for x in columnQuestion_loc]
        columnQuestion_loc = ['B' if x=='2'  else x for x in columnQuestion_loc]
        columnQuestion_loc = ['C' if x=='3'  else x for x in columnQuestion_loc]
        columnQuestion_loc = ['D' if x=='4'  else x for x in columnQuestion_loc]
        if numAlternatives_loc == 5:        
            columnQuestion_loc = ['E' if x=='5'  else x for x in columnQuestion_loc]
            columnQuestion_loc = ['X' if x=='6'  else x for x in columnQuestion_loc]
        else:
            columnQuestion_loc = ['X' if x=='5'  else x for x in columnQuestion_loc]
            
        for permutation in range(1,numSeries_loc+1):
            numQuestionPermutations_loc = int(permutations_loc[permutation-1][question_loc-1])
            correctAnswer = correctAnswers_loc[numQuestionPermutations_loc-1]
            wrongAnswers = [x for x in alternatives_loc if x != correctAnswer]
            indicesCorrectAnswer_loc = [x for x in range(numParticipants_loc) if columnQuestion_loc[x]==correctAnswer]
            indicesWrongAnswer_loc = [x for x in range(numParticipants_loc) if (columnQuestion_loc[x] in set(wrongAnswers))]
            #correctAnswers +1
            scoreQuestionsAllPermutations_loc[permutation-1,indicesCorrectAnswer_loc,numQuestionPermutations_loc-1]+=1.0
            #wrong answers -1/(numAlternatives-1)
            scoreQuestionsAllPermutations_loc[permutation-1,indicesWrongAnswer_loc,numQuestionPermutations_loc-1]-= 1.0/(float(numAlternatives_loc)-1.0)  
            #blank answers => do nothing
        counter_alternative+=1
    return scoreQuestionsAllPermutations_loc
 
def getNumberAlternatives(sheet_loc,content_loc,permutations_loc,columnSeries_loc,scoreQuestionsIndicatedSeries_loc,alternatives_loc,blankAnswer_loc,content_colNrs_loc):
    numParticipants_loc = len(scoreQuestionsIndicatedSeries_loc)
    numQuestions_loc = len(scoreQuestionsIndicatedSeries_loc[0])
    numAlternatives_loc = len(alternatives_loc)
 
    numQuestionsAlternatives_loc = numpy.zeros(numQuestions_loc*(numAlternatives_loc+1)) # + 1 for blank answer
    numQuestionsAlternatives_loc = numQuestionsAlternatives_loc.reshape(numQuestions_loc,numAlternatives_loc+1)
    #print numQuestionsAlternatives_loc.shape

    #number of alternatives per question
    #loop over question
    for question_loc in range(1,numQuestions_loc+1):
        #print "----------------------"
        #print "question " + str(question_loc)
        name_question_serie1 = "Vraag" + str(question_loc)       
        colNr_loc = content_colNrs_loc[content_loc.index(name_question_serie1)]
        #get the answers for the participants (so skip for row with name of first row)
        columnQuestion_loc=sheet_loc.col_values(colNr_loc,1,numParticipants_loc+1)
        #TODO: replace with matrixAnswers
        # replace OMR output 1, 2, 3 , 4 , 5 , 6 with A, B, C, D, E, X
        #columnQuestion_loc = map(lambda x: "A" if x=="1" else x, columnQuestion_loc)
        #columnQuestion_loc = map(lambda x: "B" if x=="2" else x, columnQuestion_loc)
        #columnQuestion_loc = map(lambda x: "C" if x=="3" else x, columnQuestion_loc)
        #columnQuestion_loc = map(lambda x: "D" if x=="4" else x, columnQuestion_loc)
        columnQuestion_loc = ['A' if x=='1'  else x for x in columnQuestion_loc]
        columnQuestion_loc = ['B' if x=='2'  else x for x in columnQuestion_loc]
        columnQuestion_loc = ['C' if x=='3'  else x for x in columnQuestion_loc]
        columnQuestion_loc = ['D' if x=='4'  else x for x in columnQuestion_loc]        
        if numAlternatives_loc == 5:        
            #columnQuestion_loc = map(lambda x: "E" if x=="5" else x, columnQuestion_loc)
            #columnQuestion_loc = map(lambda x: "X" if x=="6" else x, columnQuestion_loc)   
            columnQuestion_loc = ['E' if x=='5'  else x for x in columnQuestion_loc]
            columnQuestion_loc = ['X' if x=='6'  else x for x in columnQuestion_loc]
        else:
            #columnQuestion_loc = map(lambda x: "X" if x=="5" else x, columnQuestion_loc)
            columnQuestion_loc = ['X' if x=='5'  else x for x in columnQuestion_loc]
            

        counter_alternative = 0;
        for alternative_loc in alternatives_loc+[blankAnswer_loc]:           
            for permutation in range(1,len(permutations_loc)+1):
                #print permutation
                indicesPermutation =  [x for x in range(len(columnSeries_loc)) if columnSeries_loc[x]==permutation]                
                #print indicesPermutation
                numQuestionPermutations_loc = int(permutations_loc[permutation-1][question_loc-1])
                #print numQuestionPermutations_loc
                indicesAlternative_loc = [x for x in indicesPermutation if columnQuestion_loc[x]==alternative_loc]
                #print indicesAlternative_loc
                #print len(indicesAlternative_loc)
                #print numQuestionsAlternatives_loc[numQuestionPermutations_loc,counter_alternative]
                numQuestionsAlternatives_loc[numQuestionPermutations_loc-1,counter_alternative]+=len(indicesAlternative_loc)                 
            counter_alternative+=1
    return numQuestionsAlternatives_loc
   
def getScoreQuestionsIndicatedSeries(scoreQuestionsAllPermutations_loc,columnSeries_loc,numAlternatives_loc):
    #print "entered getScoreQuestionsIndicatedSeries"    
    numParticipants_loc = len(scoreQuestionsAllPermutations_loc[0])
    numQuestions_loc = len(scoreQuestionsAllPermutations_loc[0][0])       
    numSeries = len(scoreQuestionsAllPermutations_loc)
    #print numParticipants_loc
    #print numQuestions_loc
    #print(numSeries)
    scoreQuestionsIndicatedSeries_loc= numpy.zeros((numParticipants_loc,numQuestions_loc))
    numberCorrectAnswers_loc = numpy.zeros((numParticipants_loc))
    numberWrongAnswers_loc = numpy.zeros((numParticipants_loc))
    numberBlankAnswers_loc = numpy.zeros((numParticipants_loc))
    for participant in range(numParticipants_loc):
        
        serieIndicated = int(columnSeries_loc[participant])
        #print(serieIndicated)
        if(serieIndicated>numSeries):
            print("Error: indicated series of participant number " + str(participant) + " is higher than the number of series")
        
        #print(serieIndicated)
        scoreQuestionsIndicatedSeries_loc[participant,:] = scoreQuestionsAllPermutations_loc[serieIndicated-1,participant,:]
       # print("test")
        correctAnswers_loc = [x for x in range(numQuestions_loc) if scoreQuestionsIndicatedSeries_loc[participant,x]==1.0]
        wrongAnswers_loc = [x for x in range(numQuestions_loc) if scoreQuestionsIndicatedSeries_loc[participant,x]==-1.0/(float(numAlternatives_loc)-1.0)]
        blankAnswers_loc = [x for x in range(numQuestions_loc) if scoreQuestionsIndicatedSeries_loc[participant,x]==0.0]
       # print(correctAnswer_loc)
        numberCorrectAnswers_loc[participant] = len(correctAnswers_loc)
        numberWrongAnswers_loc[participant] = len(wrongAnswers_loc)
        numberBlankAnswers_loc[participant] = len(blankAnswers_loc)
        #print(numberCorrectAnswers_loc[participant])
        
    averageScoreQuestions_loc = scoreQuestionsIndicatedSeries_loc.sum(axis=0)/float(numParticipants_loc)    
    return scoreQuestionsIndicatedSeries_loc, averageScoreQuestions_loc, numberCorrectAnswers_loc, numberWrongAnswers_loc, numberBlankAnswers_loc

    
def getOverallStatistics(scoreQuestionsIndicatedSeries_loc,maxTotalScore_loc): 
    #print("entered getOverallStatistics")
    numParticipants_loc = len(scoreQuestionsIndicatedSeries_loc)
    #print(numParticipants_loc)
    numQuestions_loc = len(scoreQuestionsIndicatedSeries_loc[0])       
    #To calculate the total score only use the score for the series indicated by the student
    totalScore_loc = scoreQuestionsIndicatedSeries_loc.sum(axis=1)/numQuestions_loc*maxTotalScore_loc
    #print(totalScore_loc)
    # set negative scores to 0
    totalScore_loc[totalScore_loc < 0]=0
    totalScore_loc = round2(totalScore_loc)
    #print("size total score="+str(len(totalScore_loc)))
    #print totalScore
    averageScore_loc = numpy.average(totalScore_loc)
    #print("average= " + str(averageScore_loc))
    medianScore_loc = numpy.median(totalScore_loc)
    #print("median= " + str(medianScore_loc))
    standardDeviation_loc = numpy.std(totalScore_loc)
    percentagePass_loc = 100*sum(score>= maxTotalScore_loc/2.0 for score in totalScore_loc)/float(numParticipants_loc)    
    return totalScore_loc, averageScore_loc, medianScore_loc, standardDeviation_loc, percentagePass_loc
    
def getOverallStatisticsDifferentSeries(totalScoreDifferentPermutations_loc,scoreQuestionsIndicatedSeries_loc, columnSeries_loc,maxTotalScore_loc):
    #print "entered getScoreQuestionsIndicatedSeries"    
    numParticipants_loc = len(totalScoreDifferentPermutations_loc)      
    numSeries_loc = len(totalScoreDifferentPermutations_loc[0]) 
    numParticipantsSeries_loc = numpy.zeros(numSeries_loc)    
    numQuestions_loc =  len(scoreQuestionsIndicatedSeries_loc[0])
    averageScore_loc = numpy.zeros(numSeries_loc)
    medianScore_loc = numpy.zeros(numSeries_loc)
    standardDeviation_loc = numpy.zeros(numSeries_loc)    
    percentagePass_loc = numpy.zeros(numSeries_loc)
    averageScoreQuestionsDifferentSeries_loc = numpy.zeros(numQuestions_loc* numSeries_loc)
    averageScoreQuestionsDifferentSeries_loc = averageScoreQuestionsDifferentSeries_loc.reshape(numQuestions_loc, numSeries_loc)
    
    for serie in range(1,numSeries_loc+1):
        indicesSerie_loc = [x for x in range(0,numParticipants_loc) if columnSeries_loc[x]==serie]
        totalScoreSerie_loc = [totalScoreDifferentPermutations_loc[i,serie-1] for i in indicesSerie_loc]
        numParticipantsSeries_loc[serie-1] = len(totalScoreSerie_loc)
        averageScore_loc[serie-1] = numpy.average(totalScoreSerie_loc)
        medianScore_loc[serie-1] = numpy.median(totalScoreSerie_loc)
        standardDeviation_loc[serie-1] = numpy.std(totalScoreSerie_loc)
        #print totalScoreSerie_loc
        if numParticipantsSeries_loc[serie-1] == 0:
            percentagePass_loc[serie-1] = 0;
        else:
            percentagePass_loc[serie-1] = 100* sum(score>= maxTotalScore_loc/2.0 for score in totalScoreSerie_loc)/float(numParticipantsSeries_loc[serie-1]) 
        averageScoreQuestionsDifferentSeries_loc[:,serie-1] =  numpy.average(scoreQuestionsIndicatedSeries_loc[indicesSerie_loc,:],0)
    return numParticipantsSeries_loc, averageScore_loc, medianScore_loc, standardDeviation_loc, percentagePass_loc, averageScoreQuestionsDifferentSeries_loc

def calculateTotalScoreDifferentPermutations(scoreQuestionsAllPermutations_loc,maxTotalScore_loc):
    numSeries_loc = len(scoreQuestionsAllPermutations_loc)
    numParticipants_loc = len(scoreQuestionsAllPermutations_loc[0])
    numQuestions_loc = len(scoreQuestionsAllPermutations_loc[0][0])
    totalScorePermutations_loc = numpy.zeros((numParticipants_loc,numSeries_loc))    
    for serie in range(1,numSeries_loc+1):
        totalScore_temp = scoreQuestionsAllPermutations_loc[serie-1].sum(axis=1)/numQuestions_loc*maxTotalScore_loc
        totalScore_temp[totalScore_temp < 0]=0
        totalScore_temp = round2(totalScore_temp)
        totalScorePermutations_loc[:,serie-1] = totalScore_temp
    return totalScorePermutations_loc

def calculateUpperLowerStatistics(matrixAnswers_loc,content_loc,columnSeries_loc,totalScore_loc,scoreQuestionsIndicatedSeries_loc,correctAnswers_loc,alternatives_loc,blankAnswer_loc,content_colNrs_loc,permutations_loc):
    orderedDeelnemers_loc = sorted(range(len(totalScore_loc)),key=totalScore_loc.__getitem__) 
    numParticipants_loc = len(orderedDeelnemers_loc)
    #numQuestions_loc = len(scoreQuestionsIndicatedSeries_loc[0])
    #print numQuestions_loc
    #numAlternatives_loc = len(alternatives_loc)
    third_loc = int(numpy.ceil(numParticipants_loc/3.0))
    indicesUpper_loc = orderedDeelnemers_loc[numParticipants_loc-third_loc:numParticipants_loc]
    indicesLower_loc = orderedDeelnemers_loc[0:third_loc]
    indicesMiddle_loc= orderedDeelnemers_loc[third_loc:numParticipants_loc-third_loc]
    numUpper_loc = len(indicesUpper_loc)
    numLower_loc = len(indicesLower_loc)
    numMiddle_loc = len(indicesMiddle_loc)
    
    totalScoreUpper_loc = totalScore_loc[indicesUpper_loc]
    totalScoreMiddle_loc = totalScore_loc[indicesMiddle_loc]
    totalScoreLower_loc = totalScore_loc[indicesLower_loc]
    if numUpper_loc ==0:
        averageScoreUpper_loc = 0
    else:
        averageScoreUpper_loc = sum(totalScoreUpper_loc)/float(numUpper_loc)
    if numMiddle_loc == 0:
        averageScoreMiddle_loc = 0
    else:
        averageScoreMiddle_loc = sum(totalScoreMiddle_loc)/float(numMiddle_loc)
    if numLower_loc == 0:    
        averageScoreLower_loc = 0 
    else:
        averageScoreLower_loc = sum(totalScoreLower_loc)/float(numLower_loc)
    scoreQuestionsUpper_loc =  scoreQuestionsIndicatedSeries_loc[indicesUpper_loc,:]
    scoreQuestionsMiddle_loc =  scoreQuestionsIndicatedSeries_loc[indicesMiddle_loc,:]
    scoreQuestionsLower_loc =  scoreQuestionsIndicatedSeries_loc[indicesLower_loc,:]
    averageScoreQuestionsUpper_loc = scoreQuestionsUpper_loc.sum(axis=0)/float(numUpper_loc)
    averageScoreQuestionsMiddle_loc = scoreQuestionsMiddle_loc.sum(axis=0)/float(numMiddle_loc)
    averageScoreQuestionsLower_loc = scoreQuestionsLower_loc.sum(axis=0)/float(numLower_loc)
    
    numQuestions_loc = len(scoreQuestionsIndicatedSeries_loc[0])
    numAlternatives_loc = len(alternatives_loc)
    numQuestionsAlternativesUpper_loc = numpy.zeros(numQuestions_loc*(numAlternatives_loc+1)) # + 1 for blank answer
    numQuestionsAlternativesMiddle_loc = numpy.zeros(numQuestions_loc*(numAlternatives_loc+1)) # + 1 for blank answer
    numQuestionsAlternativesLower_loc = numpy.zeros(numQuestions_loc*(numAlternatives_loc+1)) # + 1 for blank answer
    numQuestionsAlternativesUpper_loc = numQuestionsAlternativesUpper_loc.reshape(numQuestions_loc,numAlternatives_loc+1)
    numQuestionsAlternativesMiddle_loc = numQuestionsAlternativesMiddle_loc.reshape(numQuestions_loc,numAlternatives_loc+1)
    numQuestionsAlternativesLower_loc = numQuestionsAlternativesLower_loc.reshape(numQuestions_loc,numAlternatives_loc+1)


    #number of alternatives per question
    #loop over question
    for question_loc in range(1,numQuestions_loc+1):
        #print "----------------------"
        #print "question " + str(question_loc)
        #name_question_serie1 = "Vraag" + str(question_loc)       
        #colNr_loc = content_colNrs_loc[content_loc.index(name_question_serie1)]
        #get the answers for the participants (so skip for row with name of first row)
        #columnQuestion_loc=sheet_loc.col_values(colNr_loc,1,numParticipants_loc+1)   
        columnQuestion_loc=matrixAnswers_loc[:,question_loc-1];
        counter_alternative = 0;
        for alternative_loc in alternatives_loc+[blankAnswer_loc]:           
            for permutation in range(1,len(permutations_loc)+1):
                #print permutation
                indicesPermutation =  [x for x in range(len(columnSeries_loc)) if columnSeries_loc[x]==permutation]                
                #print indicesPermutation
                numQuestionPermutations_loc = int(permutations_loc[permutation-1][question_loc-1])
                #print numQuestionPermutations_loc
                indicesAlternative_loc = [x for x in indicesPermutation if columnQuestion_loc[x]==alternative_loc]
                indicesAlternativeUpper_loc = [x for x in common_elements(indicesAlternative_loc,indicesUpper_loc)] 
                indicesAlternativeMiddle_loc = [x for x in common_elements(indicesAlternative_loc,indicesMiddle_loc)] 
                indicesAlternativeLower_loc = [x for x in common_elements(indicesAlternative_loc,indicesLower_loc)] 
                #print indicesAlternative_loc
                #print len(indicesAlternative_loc)
                #print numQuestionsAlternatives_loc[numQuestionPermutations_loc,counter_alternative]
                numQuestionsAlternativesUpper_loc[numQuestionPermutations_loc-1,counter_alternative]+=len(indicesAlternativeUpper_loc)                 
                numQuestionsAlternativesMiddle_loc[numQuestionPermutations_loc-1,counter_alternative]+=len(indicesAlternativeMiddle_loc)                 
                numQuestionsAlternativesLower_loc[numQuestionPermutations_loc-1,counter_alternative]+=len(indicesAlternativeLower_loc)                 
            counter_alternative+=1

    return totalScoreUpper_loc, totalScoreMiddle_loc, totalScoreLower_loc, averageScoreUpper_loc, averageScoreMiddle_loc, averageScoreLower_loc, averageScoreQuestionsUpper_loc, averageScoreQuestionsMiddle_loc, averageScoreQuestionsLower_loc, numQuestionsAlternativesUpper_loc,numQuestionsAlternativesMiddle_loc,numQuestionsAlternativesLower_loc, scoreQuestionsUpper_loc, scoreQuestionsMiddle_loc, scoreQuestionsLower_loc, numUpper_loc, numMiddle_loc, numLower_loc

def getDistributionStudents(totalScore_loc,bordersDistributionStudentsLow_loc,bordersDistributionStudentsHigh_loc):
    distributionStudentsLow_loc = numpy.zeros(len(bordersDistributionStudentsLow_loc))
    distributionStudentsHigh_loc = numpy.zeros(len(bordersDistributionStudentsHigh_loc))
    counter = 0
    for score in bordersDistributionStudentsLow_loc:
        distributionStudentsLow_loc[counter] = sum(totalScore_loc<=score)
        counter+=1
    counter = 0        
    for score in bordersDistributionStudentsHigh_loc:
        distributionStudentsHigh_loc[counter] = sum(totalScore_loc>=score)          
        counter+=1
    return distributionStudentsHigh_loc,distributionStudentsLow_loc
    
def checkMatrixAnswers(matrixAnswers_loc,alternatives_loc,blankAnswer_loc):
    #print( matrixAnswers_loc)
    #print (alternatives_loc)
    #print( blankAnswer_loc)
    if False in [ e in alternatives_loc+[blankAnswer_loc]  for e in matrixAnswers_loc.reshape(-1) ]:
        print ("ERROR: The matrix of answers does not only contain the elements " + str(alternatives_loc + [blankAnswer_loc]))
        sys.exit()

def getScoreCategories(scoreQuestionsIndicatedSeries_loc,categorieQuestions_loc):
    scoreCategories_loc = []
    for categorie in set(categorieQuestions_loc):
        questionsCategorie = numpy.where(categorieQuestions_loc==categorie)[0]
        scoreCategories_loc.append(numpy.around(numpy.average(scoreQuestionsIndicatedSeries_loc[:,questionsCategorie],axis=1)*100)) 
    return numpy.vstack(scoreCategories_loc)    