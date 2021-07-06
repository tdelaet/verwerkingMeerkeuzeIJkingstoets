# -*- coding: utf-8 -*-
"""
Created on Wed May 21 14:58:24 2014

@author: tdelaet
"""

from xlwt import  easyxf
import numpy


#check if correct answer is chosen less than tresholdCorrectAnswer %
tresholdCorrectAnswer = 0.35
tresholdWrongAnswer = 0.35
tresholdUpperCorrectAnswer = 0.35

tresholdDiffUpperLowerThree= 0.5 #treshold for best discriminating question
tresholdDiffUpperLowerTwo= 0.3
tresholdDiffUpperLowerOne= 0.15

tresholdDifficultyQuestionZero = 0.9
tresholdDifficultyQuestionOne = 0.75
tresholdDifficultyQuestionTwo = 0.5
tresholdDifficultyQuestionThree = 0.25


font_bold = "font: bold on;"
font_red = "font: color red;"
font_italic = "font: italic true;"
align_horizcenter = "align: horiz center;"
align_horizright = "align: horiz right;"
align_horizleft = "align: horiz left;"
align_vertcenter = "align: vert centre;"
align_horizvertcenter = "align: vert centre, horiz center;"
border_bottom_medium = "border: bottom medium;"
border_top_medium = "border: top medium;"
border_right_medium = "border: right medium;"
border_left_medium = "border: left medium;"
border_leftbottom_medium = "border: left medium, bottom medium;"
border_lefttop_medium = "border: left medium, top medium;"
border_righttop_medium = "border: right medium, top medium;"
border_rightbottom_medium = "border: right medium, bottom medium;"
border_all_medium = "border: bottom medium, right medium, left medium, top medium;"
pattern_solid_grey = "pattern: pattern solid, fore_colour gray25;"
align_rotated =  'align: rotation 90;'

style_title = font_bold + border_all_medium + align_horizvertcenter
style_header = font_bold + border_bottom_medium + align_horizvertcenter
#"font: bold on; align: horiz center; border: bottom medium")
style_header_borderRight = style_header + border_right_medium
#style_header_borderRight = easyxf("font: bold on; align: horiz center; border: right medium ")
style_correctAnswer = pattern_solid_grey + font_italic
#style_correctAnswer = easyxf('pattern: pattern solid, fore_colour gray25; font: italic true')
style_specialAttention = font_red

#style_correctAnswer_borderRight = easyxf('pattern: pattern solid, fore_colour gray25; font: italic true;border: right medium')
#style_specialAttention_borderRight  = easyxf('font: color red;border: right medium')
#style_correctAnswerSpecialAttention = easyxf('font: color red, italic true; pattern: pattern solid, fore_colour gray25')
#style_correctAnswerSpecialAttention_borderRight = easyxf('font: color red, italic true; pattern: pattern solid, fore_colour gray25;border: right medium')

style_border_header = easyxf('border: left thick, top thick, bottom thick, right thick')
style_borderRight = easyxf( "border: right medium ")


def write_results(outputbook,outputbookperm,numQuestions,correctAnswers,alternatives,blankAnswer,
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
                  nameQuestions,classificationQuestionsMod,categoriesQuestions,
                  bordersDistributionStudentsLow,bordersDistributionStudentsHigh,distributionStudentsLow,distributionStudentsHigh
                   ):
                      
                      
    write_scoreAllPermutations(outputbookperm,'ScoreVerschillendeSeries',numParticipants,deelnemers,numQuestions,content,content_colNrs,totalScore,totalScoreDifferentPermutations,columnSeries)
    write_overallStatistics(outputbook,'GlobaleParameters',totalScore,averageScore,medianScore,standardDeviation,percentagePass,numParticipantsSeries,averageScoreSeries,medianScoreSeries,standardDeviationSeries,percentagePassSeries,maxTotalScore)
    #write_overallStatisticsDifferentPermutations(outputbook,'GlobaleParametersSeries',numParticipantsSeries,averageScoreSeries,medianScoreSeries,standardDeviationSeries,percentagePassSeries,maxTotalScore)
    write_averageScoreQuestions(outputbook,'GemiddeldeScoreVraag',numQuestions,averageScore,averageScoreUpper,averageScoreMiddle,averageScoreLower,averageScoreQuestions,averageScoreQuestionsUpper,averageScoreQuestionsMiddle,averageScoreQuestionsLower,averageScoreSeries,averageScoreQuestionsDifferentSeries,nameQuestions,categoriesQuestions)   
    write_percentageAlternativesQuestions(outputbook,"PercentageAlternatieven",numQuestions,correctAnswers,alternatives,blankAnswer,numQuestionsAlternatives,numParticipants,nameQuestions,categoriesQuestions)
    #write_numberAlternativesQuestions(outputbook,"AantalAlternatieven",numQuestions,correctAnswers,alternatives,blankAnswer,numQuestionsAlternatives,numParticipants)
    write_percentageAlternativesQuestionsUML(outputbook,"PercentageAlternatievenUML",numQuestions,correctAnswers,alternatives,blankAnswer,numQuestionsAlternativesUpper,numQuestionsAlternativesMiddle,numQuestionsAlternativesLower,numUpper,numMiddle,numLower,nameQuestions,categoriesQuestions)
    questionClassification  = write_histogramQuestions(outputbook,"HistogramVragen",numQuestions,scoreQuestionsIndicatedSeries,averageScoreQuestions,nameQuestions,classificationQuestionsMod,categoriesQuestions)
    write_distributionStudents(outputbook,"HistogramStudenten",numParticipants,bordersDistributionStudentsLow,bordersDistributionStudentsHigh,distributionStudentsLow,distributionStudentsHigh)


def write_scoreAllPermutations(outputbookperm_loc,nameSheet_loc,numParticipants_loc,deelnemers_loc, numQuestion_loc,content_loc,content_colNrs_loc,totalScore_loc,totalScoreDifferentPermutations_loc,columnSeries_loc):
    sheetC = outputbookperm_loc.add_sheet(nameSheet_loc)


    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Score deelnemers voor verschillende reeksen",style=easyxf(style_title))
    rowCounter+=1
    
    numSeries_loc = len(totalScoreDifferentPermutations_loc[0])

    #deelnemersnummers
        #print deelnemers
    sheetC.write(rowCounter, 0,"ijkID", style=easyxf(style_header + border_right_medium)) 
    rowCounter+=1
    for i in range(0,len(deelnemers_loc)):
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[i], style=easyxf(font_bold + border_right_medium)) 
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 1;
    #total score for indicated series
    sheetC.write(rowCounter,columnCounter,"aangeduide reeks ",style=easyxf(style_header+border_right_medium))
    rowCounter+=1
    for i in range(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,totalScore_loc[i],style=easyxf(border_right_medium))
        rowCounter+=1
    columnCounter+=1;
    
    #total score for different series
    for serie in range(1,numSeries_loc+1):
        rowCounter = 1;
        sheetC.write(rowCounter,columnCounter,"reeks " + str(serie),style=easyxf(style_header+font_bold))
        rowCounter+=1
        totalScoreSerie = totalScoreDifferentPermutations_loc[:,serie-1]
        for i in range(len(totalScore_loc)):
            # if the series is the same as the one indicated 
            if (serie == columnSeries_loc[i]):
                sheetC.write(rowCounter,columnCounter,totalScoreSerie[i],style=easyxf(style_correctAnswer))
            else:# the series is different fromthe one indicated 
                if (totalScoreSerie[i]>totalScore_loc[i]):  # if score on other serie than the one indicated is higer        
                    sheetC.write(rowCounter,columnCounter,totalScoreSerie[i],style=easyxf(style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,totalScoreSerie[i])
            rowCounter+=1                    
        columnCounter+=1;                    
    
def write_overallStatistics(outputbook_loc,nameSheet_loc,totalScore_loc,averageScore_loc,medianScore_loc,standardDeviation_loc,percentagePass_loc,numParticipantsSeries_loc,averageScoreSeries_loc,medianScoreSeries_loc,standardDeviationSeries_loc,percentagePassSeries_loc,maxTotalScore_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Globale statistiek",style=easyxf(style_title))
    rowCounter+=1
    
    numParticipants_loc = len(totalScore_loc)
    #print numParticipants_loc
    #column counter
    columnCounter = 0;
    rowCounter = 1 
    
    sheetC.write(rowCounter,columnCounter,"aantal deelnemers",style=easyxf(font_bold))
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,numParticipants_loc)
    rowCounter+=1
    
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"gemiddelde score ",style=easyxf(font_bold))
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,round(averageScore_loc,2))
    rowCounter+=1
    
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"mediaan ",style=easyxf(font_bold))
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,round(medianScore_loc,2))
    rowCounter+=1

    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"standaard deviatie",style=easyxf(font_bold))
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,round(standardDeviation_loc,2))
    rowCounter+=1
        
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"% geslaagd ",style=easyxf(font_bold))
    columnCounter+=1 
    #print totalScore_loc
    sheetC.write(rowCounter,columnCounter,round(percentagePass_loc,2))
    
    rowCounter+=5
    columnCounter = 0
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Globale statistiek verschillende reeksen",style=easyxf(style_title))
    rowCounter+=1
    
    numSeries = len(numParticipantsSeries_loc)
    #print numParticipants_loc
    #column counter
    
    for serie in range(numSeries):
        columnCounter = 0;        
        sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+1,"serie " + str(serie+1),style=easyxf(style_header))
        rowCounter+=1
        
        sheetC.write(rowCounter,columnCounter,"aantal deelnemers",style=easyxf(font_bold))
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,numParticipantsSeries_loc[serie]) 
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"gemiddelde score ",style=easyxf(font_bold))
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,round(averageScoreSeries_loc[serie],2))
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"mediaan ",style=easyxf(font_bold))
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,round(medianScoreSeries_loc[serie],2))
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"standaard deviatie ",style=easyxf(font_bold))
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,round(standardDeviationSeries_loc[serie],2))
        rowCounter+=1
                
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"% geslaagd ",style=easyxf(font_bold))
        columnCounter+=1 
        #print totalScore_loc
        sheetC.write(rowCounter,columnCounter,round(percentagePassSeries_loc[serie],2))
        
        rowCounter+=1
        rowCounter+=1
      
def write_averageScoreQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,averageScore_loc,averageScoreUpper_loc,averageScoreMiddle_loc,averageScoreLower_loc,averageScoreQuestions_loc,averageScoreQuestionsUpper_loc,averageScoreQuestionsMiddle_loc,averageScoreQuestionsLower_loc,averageScoreSeries_loc,averageScoreQuestionsDifferentSeries_loc,nameQuestions_loc,categoriesQuestions_loc):
    numSeries = len(averageScoreQuestionsDifferentSeries_loc[0])  
    sheetC = outputbook_loc.add_sheet('GemScorePerVraag')
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Gemiddelde score per vraag",style=easyxf(style_title))
    rowCounter+=1
    
    #column counter
    columnCounter = 0; 


    #write all/upper/middle/lower on top
    columnCounter = 1
    sheetC.write(rowCounter+1,columnCounter,"all",style=easyxf(style_header + border_right_medium))
    columnCounter+=1
    
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+2," UML",style=easyxf(style_title))
    sheetC.write_merge(rowCounter,rowCounter,columnCounter+3,columnCounter+6," reeksen",style=easyxf(style_title))
    rowCounter+=1

    sheetC.write(rowCounter,columnCounter,"upper",style=easyxf(style_header)) 
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"middle",style=easyxf(style_header)) 
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"lower",style=easyxf(style_header+ border_right_medium)) 
    columnCounter+=1

    for serie in range(0,numSeries): #TODO: numseries
        sheetC.write(rowCounter,columnCounter,"reeks " + str(serie+1) ,style=easyxf(style_header))
        columnCounter+=1
        
    sheetC.write(rowCounter,columnCounter,"ID vraag",style=easyxf(style_header+ border_all_medium)) 
    columnCounter+=1    
    
    sheetC.write(rowCounter,columnCounter,"categorie vraag",style=easyxf(style_header+ border_all_medium)) 
    columnCounter+=1    
    
    rowCounter+=1    
    columnCounter=0
    
    for question in range(1,numQuestions_loc+1):
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf( font_bold+border_right_medium))
        columnCounter+=1        
        if averageScoreQuestions_loc[question-1]<0:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],3),style=easyxf(style_specialAttention + border_right_medium)        )
        else:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],3),style=easyxf(border_right_medium))                
        columnCounter+=1 
        if averageScoreQuestionsUpper_loc[question-1]<=averageScoreQuestionsLower_loc[question-1] or averageScoreQuestionsUpper_loc[question-1]<=averageScoreQuestionsMiddle_loc[question-1]:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsUpper_loc[question-1],3),style=easyxf(style_specialAttention))
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsMiddle_loc[question-1],3),style=easyxf(style_specialAttention)) 
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsLower_loc[question-1],3),style=easyxf(style_specialAttention+ border_right_medium))    
            columnCounter+=1
        else:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsUpper_loc[question-1],3))
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsMiddle_loc[question-1],3))
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsLower_loc[question-1],3),style=easyxf(border_right_medium))
            columnCounter+=1
        for serie in range(1,numSeries+1):
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsDifferentSeries_loc[question-1,serie-1],3))
            columnCounter+=1
        sheetC.write(rowCounter,columnCounter,nameQuestions_loc[question-1],style=easyxf( border_left_medium))    
        columnCounter+=1
        sheetC.write(rowCounter,columnCounter,categoriesQuestions_loc[question-1],style=easyxf( border_left_medium))    
        columnCounter+=1
        rowCounter+=1
        columnCounter = 0;
        
    columnCounter=0
    sheetC.write(rowCounter,columnCounter,"totaal",style=easyxf(style_header+border_all_medium))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,round(averageScore_loc,3),style=easyxf(border_righttop_medium))   
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,round(averageScoreUpper_loc,3),style=easyxf(border_top_medium))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,round(averageScoreMiddle_loc,3),style=easyxf(border_top_medium))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,round(averageScoreLower_loc,3),style=easyxf(border_righttop_medium))  
    columnCounter+=1
    for serie in range(1,numSeries+1):
        sheetC.write(rowCounter,columnCounter,round(averageScoreSeries_loc[serie-1],3),style=easyxf(border_top_medium))
        columnCounter+=1        
      

def write_percentageAlternativesQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,blankAnswer_loc,numQuestionsAlternatives_loc,numParticipants_loc,nameQuestions_loc,categoriesQuestions_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)

    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Percentage per alternatief",style=easyxf(style_title))
    rowCounter+=1
    
    #write alternative names on top
    columnCounter = 1    
    for alternative in alternatives_loc+ [blankAnswer_loc]:
        sheetC.write(rowCounter,columnCounter,alternative,style=easyxf(style_header)    )
        columnCounter+=1
        
    sheetC.write(rowCounter,columnCounter,"ID vraag",style=easyxf(style_header+ border_all_medium)) 
    columnCounter+=1  
    
    sheetC.write(rowCounter,columnCounter,"categorie vraag",style=easyxf(style_header+ border_all_medium)) 
    columnCounter+=1  
    rowCounter+=1
        
    for question in range(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(font_bold+border_right_medium))
        columnCounter+=1
        alternativeCounter = 0
        for alternative in alternatives_loc + [blankAnswer_loc]:
            if alternative == correctAnswer:
                # check if correct answer is chosen less than x %
                if (numQuestionsAlternatives_loc[question-1,alternativeCounter]/numParticipants_loc < tresholdCorrectAnswer):
                    sheetC.write(rowCounter,columnCounter,int(round(numQuestionsAlternatives_loc[question-1,alternativeCounter]/numParticipants_loc*100,0)),style = easyxf( style_correctAnswer + style_specialAttention) ) 
                else:
                    sheetC.write(rowCounter,columnCounter,int(round(numQuestionsAlternatives_loc[question-1,alternativeCounter]/numParticipants_loc*100,0)),style = easyxf( style_correctAnswer))
            else:
                # check if wrong answer is chosen more than x %
                if (numQuestionsAlternatives_loc[question-1,alternativeCounter]/numParticipants_loc > tresholdWrongAnswer):
                    sheetC.write(rowCounter,columnCounter,int(round(numQuestionsAlternatives_loc[question-1,alternativeCounter]/numParticipants_loc*100,0)),style = easyxf(style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,int(round(numQuestionsAlternatives_loc[question-1,alternativeCounter]/numParticipants_loc*100,0)))
            columnCounter+=1 
            alternativeCounter+=1
        sheetC.write(rowCounter,columnCounter,nameQuestions_loc[question-1],style=easyxf( border_left_medium))    
        columnCounter+=1    
        sheetC.write(rowCounter,columnCounter,categoriesQuestions_loc[question-1],style=easyxf( border_left_medium))    
        columnCounter+=1
        
        rowCounter+=1
        
def write_numberAlternativesQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,blankAnswer_loc,numQuestionsAlternatives_loc,numParticipants_loc,nameQuestions_loc,categoriesQuestions_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Aantal per alternatief",style=easyxf(style_title))
    rowCounter+=1
    
    #write alternative names on top
    columnCounter = 1    
    for alternative in alternatives_loc+ [blankAnswer_loc]:
        sheetC.write(rowCounter,columnCounter,alternative,style=easyxf(style_header)    )
        columnCounter+=1
        
    sheetC.write(rowCounter,columnCounter,"ID vraag",style=easyxf(style_header+ border_all_medium)) 
    columnCounter+=1 
    sheetC.write(rowCounter,columnCounter,"categorie vraag",style=easyxf(style_header+ border_all_medium)) 
    columnCounter+=1 
    rowCounter+=1
        
    for question in range(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf( font_bold+ border_right_medium)) 
        columnCounter+=1
        alternativeCounter = 0
        for alternative in alternatives_loc + [blankAnswer_loc]:
            if alternative == correctAnswer:
                if (numQuestionsAlternatives_loc[question-1,alternativeCounter]/numParticipants_loc < tresholdCorrectAnswer):
                    sheetC.write(rowCounter,columnCounter,round(numQuestionsAlternatives_loc[question-1,alternativeCounter],2),style = easyxf(style_correctAnswer + style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,round(numQuestionsAlternatives_loc[question-1,alternativeCounter],2),style = easyxf(style_correctAnswer)                    )
            else:
                if (numQuestionsAlternatives_loc[question-1,alternativeCounter]/numParticipants_loc > tresholdWrongAnswer):
                    sheetC.write(rowCounter,columnCounter,round(numQuestionsAlternatives_loc[question-1,alternativeCounter],2),style = easyxf(style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,round(numQuestionsAlternatives_loc[question-1,alternativeCounter],2))
            columnCounter+=1 
            alternativeCounter+=1
        sheetC.write(rowCounter,columnCounter,nameQuestions_loc[question-1],style=easyxf( border_left_medium))    
        columnCounter+=1    
        sheetC.write(rowCounter,columnCounter,categoriesQuestions_loc[question-1],style=easyxf(border_left_medium)) 
        columnCounter+=1          
        rowCounter+=1


def write_percentageAlternativesQuestionsUML(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,blankAnswer_loc,numQuestionsAlternativesUpper_loc,numQuestionsAlternativesMiddle_loc,numQuestionsAlternativesLower_loc,numUpper_loc,numMiddle_loc,numLower_loc,nameQuestions_loc,categoriesQuestions_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)

    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Percentage per alternatief UML",style=easyxf(style_title))
    rowCounter+=1
    
    columnCounter = 1
    for alternative in alternatives_loc+ [blankAnswer_loc]:
        sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+2,alternative,style=easyxf(style_header)    )
        sheetC.write(rowCounter+1,columnCounter,"upper",style=easyxf(style_header))
        sheetC.write(rowCounter+1,columnCounter+1,"middle",style=easyxf(style_header))
        sheetC.write(rowCounter+1,columnCounter+2,"lower",style=easyxf(style_header_borderRight + border_right_medium))
        columnCounter+=3
    sheetC.write(rowCounter,columnCounter,"onderscheidend vermogen",style=easyxf(style_header)    )
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"ID vraag",style=easyxf(style_header+ border_all_medium)) 
    columnCounter+=1 
    sheetC.write(rowCounter,columnCounter,"categorie vraag",style=easyxf(style_header+ border_all_medium)) 
    columnCounter+=1 

    rowCounter+=2
        
    for question in range(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(font_bold+border_right_medium)) 
        columnCounter+=1
        alternativeCounter = 0
        for alternative in alternatives_loc + [blankAnswer_loc]:
            if numUpper_loc == 0:
                upperPerc = 0
            else:
                upperPerc = numQuestionsAlternativesUpper_loc[question-1,alternativeCounter]/numUpper_loc  
            if numMiddle_loc == 0:
                middlePerc = 0
            else:   
                middlePerc = numQuestionsAlternativesMiddle_loc[question-1,alternativeCounter]/numMiddle_loc  
            if numLower_loc == 0:
                lowerPerc = 0
            else:
                lowerPerc = numQuestionsAlternativesLower_loc[question-1,alternativeCounter]/numLower_loc 
            if alternative == correctAnswer:
                # test of uppergroep het correcte antwoord minder  aanduidt dan lower groep
                if ( (upperPerc < lowerPerc) ):
                    sheetC.write(rowCounter,columnCounter  ,int(round(upperPerc*100,0)),style = easyxf(style_correctAnswer + style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+1,int(round(middlePerc*100,0)),style = easyxf(style_correctAnswer + style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+2,int(round(lowerPerc*100,0)),style = easyxf(style_correctAnswer+ style_specialAttention+ border_right_medium)) 
                else:
                    # test of less than X percent of upper group indicates correct answer as correct one
                    if(upperPerc<tresholdUpperCorrectAnswer):                   
                        sheetC.write(rowCounter,columnCounter,int(round(upperPerc*100,0)),style = easyxf(style_correctAnswer+ style_specialAttention))
                    else:
                        sheetC.write(rowCounter,columnCounter,int(round(upperPerc*100,0)),style = easyxf(style_correctAnswer) ) 
                    sheetC.write(rowCounter,columnCounter+1,int(round(middlePerc*100,0)),style = easyxf(style_correctAnswer) )
                    sheetC.write(rowCounter,columnCounter+2,int(round(lowerPerc*100,0)),style = easyxf( style_correctAnswer + border_right_medium))
                diffUpperLower = upperPerc - lowerPerc
            else:
                # test of uppergroep een fout antwoord meer aanduidt dan lower groep or if upper group percentage is lower than fixed number
                if (upperPerc> lowerPerc):
                    sheetC.write(rowCounter,columnCounter,int(round(upperPerc*100,0)),style = easyxf(style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+1,int(round(middlePerc*100,0)),style = easyxf(style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+2,int(round(lowerPerc*100,0)),style = easyxf(style_specialAttention + border_right_medium) )
                else:
                    # test of uppergroep een fout antwoord meer aanduidt dan goed antwoord
                    if (numQuestionsAlternativesUpper_loc[question-1,alternativeCounter] > numQuestionsAlternativesUpper_loc[(question-1),alternatives_loc.index(correctAnswer)] ):
                        sheetC.write(rowCounter,columnCounter,int(round(upperPerc*100,0)),style = easyxf(style_specialAttention))
                    else:
                        sheetC.write(rowCounter,columnCounter,int(round(upperPerc*100,0)))
                    sheetC.write(rowCounter,columnCounter+1,int(round(middlePerc*100,0)))
                    sheetC.write(rowCounter,columnCounter+2,int(round(lowerPerc*100,0)),style = easyxf(border_right_medium) ) 
            columnCounter+=3
            alternativeCounter+=1
        if diffUpperLower>=tresholdDiffUpperLowerThree:
            sheetC.write(rowCounter,columnCounter,"++++",style=easyxf(align_horizvertcenter))
        elif diffUpperLower>=tresholdDiffUpperLowerTwo:
            sheetC.write(rowCounter,columnCounter,"+++",style=easyxf(align_horizvertcenter))
        elif diffUpperLower>=tresholdDiffUpperLowerOne:
            sheetC.write(rowCounter,columnCounter,"++",style=easyxf(align_horizvertcenter))
        elif diffUpperLower>0: 
            sheetC.write(rowCounter,columnCounter,"+",style=easyxf(align_horizvertcenter))
        else: 
            sheetC.write(rowCounter,columnCounter,"0",style=easyxf(align_horizvertcenter))
        columnCounter+=1
        
        sheetC.write(rowCounter,columnCounter,nameQuestions_loc[question-1],style=easyxf( border_left_medium))    
        columnCounter+=1
        sheetC.write(rowCounter,columnCounter,categoriesQuestions_loc[question-1],style=easyxf(border_left_medium+border_right_medium)) 
        columnCounter+=1                
        rowCounter+=1

     
def write_histogramQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,scoreQuestionsIndicatedSeries_loc,averageScoreQuestions_loc,nameQuestions_loc,classificationQuestionsMod_loc,categoriesQuestions_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    numParticipants  = len(scoreQuestionsIndicatedSeries_loc)   
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Histogram score vragen",style = easyxf(style_title))
    rowCounter+=1
    
    #column counter
    columnCounter = 0;
    #gemiddelde verdeling scores per vraag
    #counter=0
    possibleScores=numpy.array([-1.0/4.0,0.0,1.0,1.0+1.0/4.0]) #TODO make parameter
    columnCounter = 1
    for possibleScore in possibleScores[0:len(possibleScores)-1]:
        sheetC.write(rowCounter,columnCounter,possibleScore,style = easyxf(style_header)) 
        columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"gemiddelde",style = easyxf(style_header))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"type vraag",style = easyxf(style_header))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"voorspelde type vraag",style=easyxf(style_header+ border_left_medium)) 
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"ID vraag",style=easyxf(style_header+ border_right_medium)) 
    columnCounter+=1 
    sheetC.write(rowCounter,columnCounter,"categorie vraag",style=easyxf(style_header+ border_right_medium)) 
    columnCounter+=1 

    rowCounter+=1    
    questionClassification=[]
    matrixClassification =numpy.array([""]*15, dtype=str).reshape(3,5)
    matrixClassificationCounter = numpy.zeros(15).reshape(3,5)
    for question in range(1,numQuestions_loc+1):
        columnCounter=0
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style = easyxf(font_bold+border_right_medium) )
        hist,bins = numpy.histogram(scoreQuestionsIndicatedSeries_loc[:,question-1],bins=possibleScores-1.0/6.0)
        columnCounter+=1    
        for n in hist:        
            if (hist[0]>hist[len(hist)-1] or hist[0]+hist[1]>hist[len(hist)-1]+hist[len(hist)-2]): #more confident in wrong answer than confident in correct answer
                sheetC.write(rowCounter,columnCounter,str(n),style = easyxf(style_specialAttention))
            else:
                #print (rowCounter)
                #print (columnCounter)
                #print (n)
                sheetC.write(rowCounter,columnCounter,str(n))
            columnCounter+=1
        if averageScoreQuestions_loc[question-1]<0:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],2),style = easyxf(style_specialAttention)        )
        else:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],2)) 
        columnCounter+=1
        correctPerc= float(hist[len(hist)-1])/numParticipants
        if correctPerc > tresholdDifficultyQuestionZero:
            sheetC.write(rowCounter,columnCounter,"0",style = easyxf(style_specialAttention+align_horizvertcenter+border_left_medium) ) 
            questionClassification.append("0")
            colClass = 0
        elif correctPerc > tresholdDifficultyQuestionOne:
            sheetC.write(rowCounter,columnCounter,"*",style=easyxf(align_horizvertcenter+border_left_medium)) 
            questionClassification.append("*")
            colClass = 1
        elif correctPerc > tresholdDifficultyQuestionTwo: 
            sheetC.write(rowCounter,columnCounter,"**",style=easyxf(align_horizvertcenter+border_left_medium))    
            questionClassification.append("**")
            colClass = 2
        elif correctPerc > tresholdDifficultyQuestionThree: 
            sheetC.write(rowCounter,columnCounter,"***",style=easyxf(align_horizvertcenter+border_left_medium))
            questionClassification.append("***")
            colClass = 3
        else: 
            sheetC.write(rowCounter,columnCounter,"****",style=easyxf(style_specialAttention+align_horizvertcenter+border_left_medium)    )
            questionClassification.append("****")
            colClass = 4
        columnCounter+=1
        sheetC.write(rowCounter,columnCounter,classificationQuestionsMod_loc[question-1],style=easyxf( border_right_medium))    
        columnCounter+=1        
        sheetC.write(rowCounter,columnCounter,nameQuestions_loc[question-1],style=easyxf( border_right_medium))    
        columnCounter+=1
        sheetC.write(rowCounter,columnCounter,categoriesQuestions_loc[question-1],style=easyxf(border_right_medium)) 
        columnCounter+=1        

        rowClass = len(classificationQuestionsMod_loc[question-1])-1       
        matrixClassification[rowClass][colClass] = matrixClassification[rowClass][colClass] + str(question) + " "
        matrixClassificationCounter[rowClass][colClass] +=1
        
        rowCounter+=1
    
    sheetC = outputbook_loc.add_sheet(nameSheet_loc+"_bis")    
    rowCounter = 0;
    columnCounter =0
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Vergelijking gemodelleerde en eigenlijke moeilijkheidsgraad vraag",style = easyxf(style_title))
    rowCounter+=1
    
    sheetC.write_merge(rowCounter,rowCounter,columnCounter+1,columnCounter+6,"eigenlijk",style = easyxf(style_title))
    rowCounter+=1
    sheetC.write_merge(rowCounter,rowCounter+3,columnCounter,columnCounter,"gemodelleerd",style = easyxf(align_rotated+border_all_medium+font_bold))

    columnCounter+=2
    sheetC.write(rowCounter,columnCounter,"0",style=easyxf(style_header))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"*",style=easyxf(style_header))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"**",style=easyxf(style_header))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"***",style=easyxf(style_header))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"****",style=easyxf(style_header))
    
    rowCounter+=1
    columnCounter=1
    for row in range(matrixClassification.shape[0]):
        if row ==0:
            sheetC.write(rowCounter,columnCounter,"*",style=easyxf(style_header+border_righttop_medium))
        elif row==1:
            sheetC.write(rowCounter,columnCounter,"**",style=easyxf(style_header+border_right_medium))
        elif row==2:
            sheetC.write(rowCounter,columnCounter,"***",style=easyxf(style_header+border_right_medium))              
        columnCounter+=1
        for col in range(matrixClassification.shape[1]):
            sheetC.write(rowCounter,columnCounter,matrixClassification[row][col])
            columnCounter+=1
        columnCounter=1
        rowCounter+=1
        
    # matrix with numbers
    columnCounter=0
    rowCounter +=3;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter+1,columnCounter+6,"eigenlijk",style = easyxf(style_title))
    rowCounter+=1
    sheetC.write_merge(rowCounter,rowCounter+3,columnCounter,columnCounter,"gemodelleerd",style = easyxf(align_rotated+border_all_medium+font_bold))

    columnCounter+=2
    sheetC.write(rowCounter,columnCounter,"0",style=easyxf(style_header))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"*",style=easyxf(style_header))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"**",style=easyxf(style_header))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"***",style=easyxf(style_header))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"****",style=easyxf(style_header))
    
    rowCounter+=1
    columnCounter=1
    for row in range(matrixClassification.shape[0]):
        if row ==0:
            sheetC.write(rowCounter,columnCounter,"*",style=easyxf(style_header+border_righttop_medium))
        elif row==1:
            sheetC.write(rowCounter,columnCounter,"**",style=easyxf(style_header+border_right_medium))
        elif row==2:
            sheetC.write(rowCounter,columnCounter,"***",style=easyxf(style_header+border_right_medium))              
        columnCounter+=1
        for col in range(matrixClassification.shape[1]):
            sheetC.write(rowCounter,columnCounter,matrixClassificationCounter[row][col])
            columnCounter+=1
        columnCounter=1
        rowCounter+=1
                
        
    return questionClassification

def write_scoreStudents(outputbook_loc,nameSheet_loc,permutations_loc,numParticipants_loc,deelnemers_loc, numQuestions_loc,numAlternatives_loc,content_loc,content_colNrs_loc,totalScore_loc,scoreQuestionsIndicatedSeries_loc,columnSeries_loc,matrixAnswers,numberCorrectAnswers_loc,numberWrongAnswers_loc,numberBlankAnswers_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)

    columnCounter = 0;
    rowCounter = 0;
    
    #deelnemersnummers
    sheetC.write(rowCounter, 0,"ijkID", style=easyxf(style_header+ border_right_medium) ) 
    rowCounter+=1
    for i in range(0,len(deelnemers_loc)):
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[i], style=easyxf(style_header + border_right_medium))
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 0;
    #total score for indicated series
    sheetC.write(rowCounter,columnCounter,"totale score",style=easyxf(style_header))
    rowCounter+=1
    for i in range(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,totalScore_loc[i])
        rowCounter+=1
    columnCounter+=1;
    

    rowCounter = 0;
    #indicated series
    sheetC.write(rowCounter,columnCounter,"reeks",style=easyxf(style_header)) 
    rowCounter+=1
    for i in range(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,columnSeries_loc[i])
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 0;
    #number of correct Answers
    sheetC.write(rowCounter,columnCounter,"aantal juist",style=easyxf(style_header)) 
    rowCounter+=1
    for i in range(len(numberCorrectAnswers_loc)):
        sheetC.write(rowCounter,columnCounter,numberCorrectAnswers_loc[i])
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 0;
    #number of wrong Answers
    sheetC.write(rowCounter,columnCounter,"aantal fout",style=easyxf(style_header)) 
    rowCounter+=1
    for i in range(len(numberWrongAnswers_loc)):
        sheetC.write(rowCounter,columnCounter,numberWrongAnswers_loc[i])
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 0;
    #number of blank Answers
    sheetC.write(rowCounter,columnCounter,"aantal blanco",style=easyxf(style_header)) 
    rowCounter+=1
    for i in range(len(numberBlankAnswers_loc)):
        sheetC.write(rowCounter,columnCounter,numberBlankAnswers_loc[i])
        rowCounter+=1
    columnCounter+=1;
    
    #score for different questions
    #beware scores are stored per question without the permutation; 
    #so for the student the scores have to be back-permutated to the order they got
    rowCounter = 0;
    #write heading    
    columnCounterHeader = columnCounter
    for question in range(1,numQuestions_loc+1):
        sheetC.write(rowCounter,columnCounterHeader,"score vraag " + str(question),style=easyxf(style_header))
        columnCounterHeader+=1
        
    columnCounterScoreQuestions = columnCounter   
    rowCounter = 1;  
    for participant in range(len(totalScore_loc)): # loop over participants
        columnCounter = columnCounterScoreQuestions;
        score = scoreQuestionsIndicatedSeries_loc[participant,:]
        
        sorted_score = [score[int(i-1)] for i in permutations_loc[int(columnSeries_loc[participant]-1)]]
        for question in range(1,numQuestions_loc+1):
            sheetC.write(rowCounter,columnCounter,sorted_score[question-1])
            columnCounter+=1
        rowCounter+=1            
 
    #answer for different questions and alternatives
    for question in range(1,numQuestions_loc+1):
            rowCounter=0
            sheetC.write(rowCounter,columnCounter,"antwoord vraag " + str(question),style=easyxf(style_header) ) 
            #print "antwoord vraag " + str(question) + str(alternative)
            answer = matrixAnswers[:,question-1]
            rowCounter+=1
            #print 
            for i in range(len(totalScore_loc)):
                sheetC.write(rowCounter,columnCounter,answer[i])
                rowCounter+=1                    
            columnCounter+=1;    
            
def write_scoreCategoriesStudents(outputbook_loc,nameSheet_loc,deelnemers_loc,totalScore_loc, categoriesQuestions_loc, scoreCategories_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)

    columnCounter = 0;
    rowCounter = 0;
    
    #deelnemersnummers
    sheetC.write(rowCounter, columnCounter,"ijkID", style=easyxf(style_header+ border_right_medium) )
    columnCounter+=1
    sheetC.write(rowCounter, columnCounter,"totale score", style=easyxf(style_header+ border_right_medium) )
    rowCounter+=1
    columnCounter=0
    
    for i in range(0,len(deelnemers_loc)):
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[i], style=easyxf(style_header + border_right_medium))
        sheetC.write(rowCounter,columnCounter+1,totalScore_loc[i], style=easyxf(border_right_medium))
        rowCounter+=1
    columnCounter+=2;
    
    columnCounterCat = columnCounter;
    rowCounter = 0;
    
    for categorie in set(categoriesQuestions_loc):
        sheetC.write(rowCounter,columnCounter,categorie,style=easyxf(style_header))
        columnCounter+=1
    rowCounter+=1

    columnCounter=columnCounterCat;
    for deelnemer in range(scoreCategories_loc.shape[1]):
        for categorie in range(len(set(categoriesQuestions_loc))):
            sheetC.write(rowCounter,columnCounter,scoreCategories_loc[categorie][deelnemer])
            columnCounter+=1
        rowCounter+=1
        columnCounter=columnCounterCat
    
#def write_overallStatisticsInstellingen(outputbook_loc,nameSheet_loc,numParticipants_loc,totalScore_loc,averageScore_loc,medianScore_loc,standardDeviation_loc,percentagePass_loc,numParticipantsSeries_loc,averageScoreSeries_loc,medianScoreSeries_loc,standardDeviationSeries_loc,percentagePassSeries_loc,maxTotalScore_loc):
def write_overallStatisticsInstellingen(outputbook_loc,nameSheet_loc,instellingen_loc,numParticipants_tot_loc,numParticipants_stacked_tot_loc,averageScore_tot_loc,averageScore_stacked_tot_loc,medianScore_tot_loc,medianScore_stacked_tot_loc,standardDeviation_tot_loc,standardDeviation_stacked_tot_loc,percentagePass_tot_loc,percentagePass_stacked_tot_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Globale statistiek",style=easyxf(style_title))
    rowCounter+=1
    

    #print numParticipants_loc
    #column counter
    columnCounter = 0;
    rowCounter = 1 
    
    sheetC.write(rowCounter,columnCounter,"aantal deelnemers",style=easyxf(font_bold))
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,str(numParticipants_tot_loc))
    rowCounter+=1
    
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"gemiddelde score ",style=easyxf(font_bold))
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,round(averageScore_tot_loc,2))
    rowCounter+=1
    
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"mediaan ",style=easyxf(font_bold))
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,round(medianScore_tot_loc,2))
    rowCounter+=1

    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"standaard deviatie",style=easyxf(font_bold))
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,round(standardDeviation_tot_loc,2))
    rowCounter+=1
        
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"% geslaagd ",style=easyxf(font_bold))
    columnCounter+=1 
    #print totalScore_loc
    sheetC.write(rowCounter,columnCounter,round(percentagePass_tot_loc,2))
    
    rowCounter+=5
    columnCounter = 0
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Globale statistiek verschillende reeksen",style=easyxf(style_title))
    rowCounter+=1
    
    numInstellingen = len(numParticipants_stacked_tot_loc)
    #print numParticipants_loc
    #column counter
    
    for instelling in range(numInstellingen):
        columnCounter = 0;        
        sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+1,"instelling " + instellingen_loc[instelling],style=easyxf(style_header))
        rowCounter+=1
        
        sheetC.write(rowCounter,columnCounter,"aantal deelnemers",style=easyxf(font_bold))
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,str(numParticipants_stacked_tot_loc[instelling][0]) )
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"gemiddelde score ",style=easyxf(font_bold))
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,round(averageScore_stacked_tot_loc[instelling][0],2))
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"mediaan ",style=easyxf(font_bold))
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,round(medianScore_stacked_tot_loc[instelling][0],2))
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"standaard deviatie ",style=easyxf(font_bold))
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,round(standardDeviation_stacked_tot_loc[instelling][0],2))
        rowCounter+=1
                
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"% geslaagd ",style=easyxf(font_bold))
        columnCounter+=1 
        #print totalScore_loc
        sheetC.write(rowCounter,columnCounter,round(percentagePass_stacked_tot_loc[instelling][0],2))
        
        rowCounter+=1
        rowCounter+=1     
        
def write_distributionStudents(outputbook_loc,nameSheet_loc,numParticpants_loc,bordersDistributionStudentsLow_loc,bordersDistributionStudentsHigh_loc,distributionStudentsLow_loc,distributionStudentsHigh_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Histogram studenten",style=easyxf(style_title))
    rowCounter+=2
    columnCounter=0
        
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+4,"percentage met score <= ",style=easyxf(style_title))
    rowCounter+=1
    columnCounter=0
    
    sheetC.write(rowCounter,columnCounter,"score ",style=easyxf(font_bold+border_bottom_medium+border_right_medium))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"percentage ",style=easyxf(font_bold+border_bottom_medium+border_right_medium))
    columnCounter+=1
    
    rowCounter+=1
    columnCounter=0
    counter = 0
    for score in bordersDistributionStudentsLow_loc:
        columnCounter=0
        sheetC.write(rowCounter,columnCounter,score,style=easyxf(border_right_medium+font_bold))
        columnCounter+=1
        sheetC.write(rowCounter,columnCounter,round(float(distributionStudentsLow_loc[counter])/float(numParticpants_loc)*100,1))
        columnCounter+=1  
        
        rowCounter+=1
        counter+=1
        
    rowCounter+=2
    columnCounter=0
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+4,"percentage met score >= ",style=easyxf(style_title))
    rowCounter+=1
    columnCounter=0

    sheetC.write(rowCounter,columnCounter,"score ",style=easyxf(font_bold+border_bottom_medium+border_right_medium))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"percentage ",style=easyxf(font_bold+border_bottom_medium+border_right_medium))
    columnCounter+=1
    
    rowCounter+=1
    columnCounter=0
    counter = 0
    for score in bordersDistributionStudentsHigh_loc:
        columnCounter=0
        sheetC.write(rowCounter,columnCounter,score,style=easyxf(border_right_medium+font_bold))
        columnCounter+=1
        sheetC.write(rowCounter,columnCounter,round(float(distributionStudentsHigh_loc[counter])/float(numParticpants_loc)*100,1))
        columnCounter+=1  
        
        rowCounter+=1
        counter+=1        
        
def write_feedbackStudents(outputbook_loc,permutations_loc,numParticipants_loc,deelnemers_loc, numQuestions_loc,alternatives_loc,numAlternatives_loc,content_loc,content_colNrs_loc,totalScore_loc,scoreQuestionsIndicatedSeries_loc,columnSeries_loc,matrixAnswers,categorieQuestions_loc,scoreCategories_loc,
                           averageScoreQuestions_tot_loc,averageScoreQuestionsUpper_tot_loc,averageScoreQuestionsMiddle_tot_loc,averageScoreQuestionsLower_tot_loc
                           ,correctAnswers_loc, numQuestionsAlternatives_loc):
    
    orderedParticipants = sorted(range(len(deelnemers_loc)), key=lambda k: deelnemers_loc[k])
    
    for participant in orderedParticipants: # range(len(totalScore_loc)):
        sheetC = outputbook_loc.add_sheet(str(int(deelnemers_loc[participant])))

        columnCounter = 0;
        rowCounter = 0;
        
        #deelnemersnummers
        sheetC.write(rowCounter, 0,"ijkID: ", style=easyxf(font_bold + border_lefttop_medium) ) 
        columnCounter +=1
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[participant], style=easyxf(font_bold + border_righttop_medium))
        
        rowCounter+=1
        columnCounter = 0
        
        #total score for indicated series
        sheetC.write(rowCounter,columnCounter,"score",style=easyxf(border_left_medium + font_bold))
        columnCounter+=1
        sheetC.write(rowCounter,columnCounter,totalScore_loc[participant],style=easyxf(border_right_medium))
        
        rowCounter+=1;
        columnCounter = 0;
        
        #indicated series
        sheetC.write(rowCounter,columnCounter,"reeks",style=easyxf(font_bold + border_leftbottom_medium)) 
        columnCounter+=1
        sheetC.write(rowCounter,columnCounter,columnSeries_loc[participant],style=easyxf(border_rightbottom_medium))
        rowCounter+=1;
        
        rowCounter +=2; 
        rowBegin2 = rowCounter
        
        columnCounter=0
        counterCategorie = 0
        sheetC.write(rowCounter,columnCounter,"",style=easyxf(border_bottom_medium))
        columnCounter+=1
        sheetC.write(rowCounter,columnCounter,"",style=easyxf(border_bottom_medium))
        rowCounter+=1
        columnCounter=0
        for categorie in set(categorieQuestions_loc):
            sheetC.write(rowCounter,columnCounter,categorie,style=easyxf(font_bold + border_left_medium))
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,scoreCategories_loc[counterCategorie][participant],style=easyxf(font_bold + border_right_medium))
            rowCounter+=1
            columnCounter-=1
            counterCategorie+=1
        sheetC.write(rowCounter,columnCounter,"",style=easyxf(border_top_medium))
        columnCounter+=1
        sheetC.write(rowCounter,columnCounter,"",style=easyxf(border_top_medium))
        rowCounter+=1
        
        score = scoreQuestionsIndicatedSeries_loc[participant,:]
        sorted_score = [score[int(i-1)] for i in permutations_loc[int(columnSeries_loc[participant]-1)]]
        
        numBlank = sum(score == 0)
        numCorrect = sum(score == 1.0)
        numWrong = sum(score == -1.0/(numAlternatives_loc-1))    
        
        columnOffset = 4
        columnCounter = columnOffset
        rowCounter = rowBegin2
        sheetC.write(rowCounter,columnCounter,"juist",style=easyxf(font_bold + border_lefttop_medium))
        columnCounter +=1
        sheetC.write(rowCounter,columnCounter,str(numCorrect),style=easyxf( border_righttop_medium))
        rowCounter+=1
        
        columnCounter = columnOffset
        sheetC.write(rowCounter,columnCounter,"fout",style=easyxf(font_bold+border_left_medium))
        columnCounter +=1
        sheetC.write(rowCounter,columnCounter,str(numWrong),style=easyxf(border_right_medium))
        rowCounter+=1
        
        columnCounter = columnOffset
        sheetC.write(rowCounter,columnCounter,"blanco",style=easyxf(font_bold+border_leftbottom_medium))
        columnCounter +=1
        sheetC.write(rowCounter,columnCounter,str(numBlank),style=easyxf(border_rightbottom_medium))
        rowCounter+=1
        
        rowOffset = rowCounter+len(set(categorieQuestions_loc)); 
        
        #score for different questions
        #beware scores are stored per question without the permutation; 
        #so for the student the scores have to be back-permutated to the order they got
        rowCounter = rowOffset;
        columnCounter = 0
        
        #write heading  
        sheetC.write(rowCounter,columnCounter,"vraag",style=easyxf(style_header))
        columnCounter+=1
        sheetC.write(rowCounter,columnCounter,"score",style=easyxf(style_header))
        columnCounter+=1        
        sheetC.write(rowCounter,columnCounter,"antwoord",style=easyxf(style_header))       
        columnCounter+=1       
        sheetC.write(rowCounter,columnCounter,"sleutel",style=easyxf(style_header))       
        columnCounter+=1   
        sheetC.write(rowCounter,columnCounter,"vraagnr.",style=easyxf(style_header)) 
        columnCounter+=1        
        sheetC.write(rowCounter,columnCounter,"type",style=easyxf(style_header)) 
        columnCounter+=1        
        sheetC.write(rowCounter,columnCounter,"gem.",style=easyxf(style_header))    
        columnCounter+=1        
        sheetC.write(rowCounter,columnCounter,"upper",style=easyxf(style_header))
        columnCounter+=1        
        sheetC.write(rowCounter,columnCounter,"lower",style=easyxf(style_header)) 
        columnCounter+=1        
        sheetC.write(rowCounter,columnCounter,"% juist",style=easyxf(style_header))
        columnCounter+=1        
        sheetC.write(rowCounter,columnCounter,"% blanco",style=easyxf(style_header))

        rowCounter+=1
        

        for question in range(1,numQuestions_loc+1):
            columnCounter = 0
            questionNumberSerie1 = permutations_loc[int(columnSeries_loc[participant]-1),int(question-1)]
            correctAnswer = correctAnswers_loc[int(questionNumberSerie1-1)]
            #print("test")
            #print(questionNumberSerie1)
            sheetC.write(rowCounter,columnCounter,str(question),style=easyxf(font_bold + border_right_medium + align_horizright))
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,sorted_score[int(question-1)],style=easyxf(align_horizleft))
            columnCounter+=1
            answer = matrixAnswers[participant,question-1]
            sheetC.write(rowCounter,columnCounter,answer,style=easyxf(align_horizleft))   
            columnCounter+=1   
            sheetC.write(rowCounter,columnCounter,correctAnswer,style=easyxf(align_horizleft))   
            columnCounter+=1  
            sheetC.write(rowCounter,columnCounter,questionNumberSerie1,style=easyxf(align_horizleft))   
            columnCounter+=1   
            sheetC.write(rowCounter,columnCounter,categorieQuestions_loc[int(questionNumberSerie1-1)],style=easyxf(align_horizleft))  
            columnCounter+=1;    
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_tot_loc[int(questionNumberSerie1-1)],2),style=easyxf(align_horizleft))  
            columnCounter+=1;    
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsUpper_tot_loc[int(questionNumberSerie1-1)],2),style=easyxf(align_horizleft) )
            columnCounter+=1;    
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsLower_tot_loc[int(questionNumberSerie1-1)],2),style=easyxf(align_horizleft))             
 
            percCorrect = int(round(numQuestionsAlternatives_loc[int(questionNumberSerie1-1),alternatives_loc.index(correctAnswer)]/numParticipants_loc*100,0))
            columnCounter+=1;    
            sheetC.write(rowCounter,columnCounter,percCorrect,style=easyxf(align_horizleft))
            percBlank = int(round(numQuestionsAlternatives_loc[int(questionNumberSerie1-1),numAlternatives_loc]/numParticipants_loc*100,0))
            columnCounter+=1;    
            sheetC.write(rowCounter,columnCounter,percBlank,style=easyxf(align_horizleft))            
            rowCounter+=1;  

def write_feedbackPlatform(outputFolder_loc,permutations_loc,numParticipants_loc,deelnemers_loc, numQuestions_loc,alternatives_loc,numAlternatives_loc,content_loc,content_colNrs_loc,totalScore_loc,scoreQuestionsIndicatedSeries_loc,columnSeries_loc,matrixAnswers,categorieQuestions_loc,scoreCategories_loc,
                           averageScoreQuestions_tot_loc,averageScoreQuestionsUpper_tot_loc,averageScoreQuestionsMiddle_tot_loc,averageScoreQuestionsLower_tot_loc
                           ,correctAnswers_loc, numQuestionsAlternatives_loc,blankAnswer_loc):
    fFeedback= open(outputFolder_loc + 'feedbackPlatform.csv','w')
    orderedParticipants = sorted(range(len(deelnemers_loc)), key=lambda k: deelnemers_loc[k])
    # write header
    fFeedback.write("ijkID" )
    fFeedback.write(',')
    fFeedback.write("reeks")
    fFeedback.write(',')
    fFeedback.write("score")
    fFeedback.write(',')
    #namen categoriën
    for categorie in set(categorieQuestions_loc):
        fFeedback.write("cat: " + categorie)
        fFeedback.write(',')
    fFeedback.write("aantal juist")
    fFeedback.write(',')
    fFeedback.write("aantal fout") 
    fFeedback.write(',')
    fFeedback.write("aantal blanco")
    fFeedback.write(',')
    for question in range(1,numQuestions_loc+1):
        fFeedback.write("vraag" + str(question) +": score")
        fFeedback.write(',') 
        fFeedback.write("vraag" + str(question) +": antwoord")       
        fFeedback.write(',')
        fFeedback.write("vraag" + str(question) +": sleutel")       
        fFeedback.write(',')
        fFeedback.write("vraag" + str(question) +": type") 
        fFeedback.write(',')
        fFeedback.write("vraag" + str(question) +": gem.")    
        fFeedback.write(',')
        fFeedback.write("vraag" + str(question) +": upper")
        fFeedback.write(',')
        fFeedback.write("vraag" + str(question) +": lower") 
        fFeedback.write(',')
        fFeedback.write("vraag" + str(question) + ": %juist")
        fFeedback.write(',')
        fFeedback.write("vraag" + str(question) +": %blanco")
        fFeedback.write(',')  
        for alternative in alternatives_loc:
            fFeedback.write("aantal " +alternative)
            fFeedback.write(',')  
        fFeedback.write("aantal " + blankAnswer_loc)
        fFeedback.write(',')          
    fFeedback.write('\n')
    for participant in orderedParticipants: # range(len(totalScore_loc)):
        #ijkID
        fFeedback.write(str(int(deelnemers_loc[participant])))
        fFeedback.write(',')
        
        #indicated series
        fFeedback.write(str(int(columnSeries_loc[participant])))      
        fFeedback.write(',')
        
        #total score for indicated series
        fFeedback.write(str(int(totalScore_loc[participant])))
        fFeedback.write(',')

        #score per categorie
        counterCategorie = 0;
        for categorie in set(categorieQuestions_loc):
            fFeedback.write(str(scoreCategories_loc[counterCategorie][participant]))
            fFeedback.write(',')
            counterCategorie+=1
        
        score = scoreQuestionsIndicatedSeries_loc[participant,:]
        sorted_score = [score[int(i-1)] for i in permutations_loc[int(columnSeries_loc[participant]-1)]]
        
        numBlank = sum(score == 0)
        numCorrect = sum(score == 1.0)
        numWrong = sum(score == -1.0/(numAlternatives_loc-1))    
        
       
        # aantal juist
        fFeedback.write(str(numCorrect))
        fFeedback.write(',')
        # aantal fout
        fFeedback.write(str(numWrong))
        fFeedback.write(',')
        # aantal blanco
        fFeedback.write(str(numBlank))
        fFeedback.write(',')
        
        
        
        

        #score for different questions
        #beware scores are stored per question without the permutation; 
        #so for the student the scores have to be back-permutated to the order they got
        
        for question in range(1,numQuestions_loc+1):
            #print(int(columnSeries_loc[int(participant-1)]))
            #print(int(question-1))
            #print(permutations_loc)
            questionNumberSerie1 = permutations_loc[int(columnSeries_loc[participant]-1),int(question-1)]
            correctAnswer = correctAnswers_loc[int(questionNumberSerie1-1)]

            #fFeedback.write(str(question),style=easyxf(font_bold + border_right_medium + align_horizright))
            #columnCounter+=1
            fFeedback.write(str(sorted_score[int(question-1)]))
            fFeedback.write(',')
            answer = matrixAnswers[participant,int(question-1)]
            fFeedback.write(answer)   
            fFeedback.write(',')
            fFeedback.write(correctAnswer)   
            fFeedback.write(',')
            fFeedback.write(categorieQuestions_loc[int(questionNumberSerie1-1)])  
            fFeedback.write(',')   
            fFeedback.write(str((round(averageScoreQuestions_tot_loc[int(questionNumberSerie1-1)],2))  ))
            fFeedback.write(',')
            fFeedback.write(str(int(round(averageScoreQuestionsUpper_tot_loc[int(questionNumberSerie1-1)]*100,0) )))
            fFeedback.write(',')
            fFeedback.write(str(int(round(averageScoreQuestionsLower_tot_loc[int(questionNumberSerie1-1)]*100,0))           )  )
 
            percCorrect = int(round(numQuestionsAlternatives_loc[int(questionNumberSerie1-1),alternatives_loc.index(correctAnswer)]/numParticipants_loc*100,0))
            fFeedback.write(',')    
            fFeedback.write(str(percCorrect))
            percBlank = int(round(numQuestionsAlternatives_loc[int(questionNumberSerie1-1),numAlternatives_loc]/numParticipants_loc*100,0))
            fFeedback.write(',')
            fFeedback.write(str(percBlank))  
            fFeedback.write(',')
            
            for alternative in range(0,numAlternatives_loc):
                aantalAlternative = int(round(numQuestionsAlternatives_loc[int(questionNumberSerie1-1),alternative],0))
                fFeedback.write(str(aantalAlternative))
                fFeedback.write(',')
            aantalBlank = int(round(numQuestionsAlternatives_loc[int(questionNumberSerie1-1),numAlternatives_loc],0))
            fFeedback.write(str(aantalBlank))
            fFeedback.write(',')            
        fFeedback.write('\n')  
    fFeedback.close()
            
            
def write_scoreStudentsNonPermutated(outputbook_loc,nameSheet_loc,permutations_loc,numParticipants_loc,deelnemers_loc, numQuestions_loc,numAlternatives_loc,alternatives_loc,content_loc,content_colNrs_loc,totalScore_loc,scoreQuestionsIndicatedSeries_loc,columnSeries_loc,matrixAnswers):
    
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)

    columnCounter = 0;
    rowCounter = 0;
    
    #deelnemersnummers
    sheetC.write(rowCounter, 0,"studentennummer", style=easyxf(style_header_borderRight)) 
    rowCounter+=1
    for i in range(0,len(deelnemers_loc)):
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[i], style=easyxf(style_header_borderRight))
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 0;
    #total score for indicated series
    sheetC.write(rowCounter,columnCounter,"totale score",style=easyxf(style_header))
    rowCounter+=1
    for i in range(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,totalScore_loc[i])
        rowCounter+=1
    columnCounter+=1;
    

    rowCounter = 0;
    #indicated series
    sheetC.write(rowCounter,columnCounter,"reeks",style=easyxf(style_header))
    rowCounter+=1
    for i in range(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,columnSeries_loc[i])
        rowCounter+=1
    columnCounter+=1;
    
    #score for different questions
    rowCounter = 0;
    #write heading    
    columnCounterHeader = columnCounter
    for question in range(1,numQuestions_loc+1):
        sheetC.write(rowCounter,columnCounterHeader,"score vraag " + str(question),style=easyxf(style_header))
        columnCounterHeader+=1
        
    columnCounterScoreQuestions = columnCounter   
    rowCounter = 1;  
    for participant in range(len(totalScore_loc)): # loop over participants
        columnCounter = columnCounterScoreQuestions;
        score = scoreQuestionsIndicatedSeries_loc[participant,:]
        #serie = int(columnSeries_loc[participant]-1)
        #sorted_score = [score[i-1] for i in permutations_loc[serie]]
        #find questions with weight zero
        #questionsZeroWeight = numpy.where(weightsQuestions_loc==0)
                #find questions
        #score[numpy.where(weightsQuestions_loc==0)[0]]=float('NaN')
        
        for question in range(1,numQuestions_loc+1):
            #if (question-1) in questionsZeroWeight:
            #    sheetC.write(rowCounter,columnCounter,"X")
            #else:
            #    sheetC.write(rowCounter,columnCounter,score[question-1])
            sheetC.write(rowCounter,columnCounter,score[int(question-1)])
            columnCounter+=1
        rowCounter+=1            
 
        columnCounter = columnCounterScoreQuestions+1;
        
    #answers for different questions
    #beware answers are stored per question with the permutation; 
    #so for the us the answers have to be permutated 
    #write heading    
    for question in range(1,numQuestions_loc+1):
        for alternative in alternatives_loc:
            sheetC.write(0,columnCounterHeader,"antwoord vraag " + str(question) + alternative,style=easyxf(style_header))
            columnCounterHeader+=1
    rowCounter = 1;      
    columnCounterAnswers = columnCounter
#   comment Riet 12/5/2016: onderstaande nog te debuggen    
#    for participant in range(len(totalScore_loc)): # loop over participants
#        columnCounter = columnCounterAnswers;
#        serie = int(columnSeries_loc[participant]-1)
#        answers =  matrixAnswers[participant]
#        
#        #find questions with weight zero
#        #questionsZeroWeight = numpy.where(weightsQuestions_loc==0)
#                #find questions
#        #score[numpy.where(weightsQuestions_loc==0)[0]]=float('NaN')
#        
#        for question in range(1,numQuestions_loc+1):
#            questionInSerie = numpy.where(permutations_loc[serie]==question)[0][0]+1
#            answersQuestion = answers[numAlternatives_loc*(questionInSerie-1):numAlternatives_loc*(questionInSerie-1)+numAlternatives_loc]                                  
#            for counterAlternative in range(0,numAlternatives_loc):
#                print rowCounter,columnCounter,answersQuestion[counterAlternative]
#                sheetC.write(rowCounter,columnCounter,answersQuestion[counterAlternative])
#                columnCounter+=1
#        rowCounter+=1             

def write_participantsList(outputbook_loc,nameSheet_loc,deelnemers_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)

    columnCounter = 0;
    rowCounter = 0;
    sheetC.write(rowCounter,0,"1 tot 20")
    
    rowCounter+=1
    sheetC.write(rowCounter,0,"Naam")
    sheetC.write(rowCounter,1,"Voornaam")
    sheetC.write(rowCounter,2,"Studnr")
    sheetC.write(rowCounter,10,"TOTAAL")
    
    rowCounter+=1
    columnCounter=2
    #deelnemersnummers
    for i in range(0,len(deelnemers_loc)):
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[i])
        rowCounter+=1
  