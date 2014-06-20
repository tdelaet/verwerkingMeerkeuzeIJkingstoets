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
align_vertcenter = "align: vert centre;"
align_horizvertcenter = "align: vert centre, horiz center;"
border_bottom_medium = "border: bottom medium;"
border_top_medium = "border: top medium;"
border_right_medium = "border: right medium;"
border_righttop_medium = "border: right medium, top medium;"
border_all_medium = "border: bottom medium, right medium, left medium, top medium;"
pattern_solid_grey = "pattern: pattern solid, fore_colour gray25;"

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


def write_results(outputbook,numQuestions,correctAnswers,alternatives,blankAnswer,
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
                   ):
                      
                      
    write_scoreAllPermutations(outputbook,'ScoreVerschillendeSeries',numParticipants,deelnemers,numQuestions,content,content_colNrs,totalScore,totalScoreDifferentPermutations,columnSeries)
    write_overallStatistics(outputbook,'GlobaleParameters',totalScore,averageScore,medianScore,standardDeviation,percentagePass,numParticipantsSeries,averageScoreSeries,medianScoreSeries,standardDeviationSeries,percentagePassSeries,maxTotalScore)
    #write_overallStatisticsDifferentPermutations(outputbook,'GlobaleParametersSeries',numParticipantsSeries,averageScoreSeries,medianScoreSeries,standardDeviationSeries,percentagePassSeries,maxTotalScore)
    write_averageScoreQuestions(outputbook,'GemiddeldeScoreVraag',numQuestions,averageScore,averageScoreUpper,averageScoreMiddle,averageScoreLower,averageScoreQuestions,averageScoreQuestionsUpper,averageScoreQuestionsMiddle,averageScoreQuestionsLower,averageScoreSeries,averageScoreQuestionsDifferentSeries)   
    write_percentageAlternativesQuestions(outputbook,"PercentageAlternatieven",numQuestions,correctAnswers,alternatives,blankAnswer,numQuestionsAlternatives,numParticipants)
    #write_numberAlternativesQuestions(outputbook,"AantalAlternatieven",numQuestions,correctAnswers,alternatives,blankAnswer,numQuestionsAlternatives,numParticipants)
    write_percentageAlternativesQuestionsUML(outputbook,"PercentageAlternatievenUML",numQuestions,correctAnswers,alternatives,blankAnswer,numQuestionsAlternativesUpper,numQuestionsAlternativesMiddle,numQuestionsAlternativesLower,numUpper,numMiddle,numLower)
    write_histogramQuestions(outputbook,"HistogramVragen",numQuestions,scoreQuestionsIndicatedSeries,averageScoreQuestions)



def write_scoreAllPermutations(outputbook_loc,nameSheet_loc,numParticipants_loc,deelnemers_loc, numQuestion_loc,content_loc,content_colNrs_loc,totalScore_loc,totalScoreDifferentPermutations_loc,columnSeries_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)


    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Score deelnemers voor verschillende reeksen",style=easyxf(style_title))
    rowCounter+=1
    
    numSeries_loc = len(totalScoreDifferentPermutations_loc[0])

    #deelnemersnummers
        #print deelnemers
    sheetC.write(rowCounter, 0,"ijkID", style=easyxf(style_header + border_right_medium)) 
    rowCounter+=1
    for i in xrange(0,len(deelnemers_loc)):
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[i], style=easyxf(font_bold + border_right_medium)) 
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 1;
    #total score for indicated series
    sheetC.write(rowCounter,columnCounter,"aangeduide reeks ",style=easyxf(style_header+border_right_medium))
    rowCounter+=1
    for i in xrange(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,totalScore_loc[i],style=easyxf(border_right_medium))
        rowCounter+=1
    columnCounter+=1;
    
    #total score for different series
    for serie in xrange(1,numSeries_loc+1):
        rowCounter = 1;
        sheetC.write(rowCounter,columnCounter,"reeks " + str(serie),style=easyxf(style_header+font_bold))
        rowCounter+=1
        totalScoreSerie = totalScoreDifferentPermutations_loc[:,serie-1]
        for i in xrange(len(totalScore_loc)):
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
    
    for serie in xrange(numSeries):
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
      
def write_averageScoreQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,averageScore_loc,averageScoreUpper_loc,averageScoreMiddle_loc,averageScoreLower_loc,averageScoreQuestions_loc,averageScoreQuestionsUpper_loc,averageScoreQuestionsMiddle_loc,averageScoreQuestionsLower_loc,averageScoreSeries_loc,averageScoreQuestionsDifferentSeries_loc):
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

    for serie in xrange(0,numSeries): #TODO: numseries
        sheetC.write(rowCounter,columnCounter,"reeks " + str(serie+1) ,style=easyxf(style_header))
        columnCounter+=1
    
    rowCounter+=1    
    columnCounter=0
    
    for question in xrange(1,numQuestions_loc+1):
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
        else:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsUpper_loc[question-1],3))
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsMiddle_loc[question-1],3))
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsLower_loc[question-1],3),style=easyxf(border_right_medium))
            columnCounter+=1
        for serie in xrange(1,numSeries+1):
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsDifferentSeries_loc[question-1,serie-1],3))
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
    for serie in xrange(1,numSeries+1):
        sheetC.write(rowCounter,columnCounter,round(averageScoreSeries_loc[serie-1],3),style=easyxf(border_top_medium))
        columnCounter+=1
      

def write_percentageAlternativesQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,blankAnswer_loc,numQuestionsAlternatives_loc,numParticipants_loc):
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
    rowCounter+=1
        
    for question in xrange(1,numQuestions_loc+1):
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
                    
        rowCounter+=1
        
def write_numberAlternativesQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,blankAnswer_loc,numQuestionsAlternatives_loc,numParticipants_loc):
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
    rowCounter+=1
        
    for question in xrange(1,numQuestions_loc+1):
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
        rowCounter+=1


def write_percentageAlternativesQuestionsUML(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,blankAnswer_loc,numQuestionsAlternativesUpper_loc,numQuestionsAlternativesMiddle_loc,numQuestionsAlternativesLower_loc,numUpper_loc,numMiddle_loc,numLower_loc):
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
    rowCounter+=2
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(font_bold+border_right_medium)) 
        columnCounter+=1
        alternativeCounter = 0
        for alternative in alternatives_loc + [blankAnswer_loc]:
            upperPerc = numQuestionsAlternativesUpper_loc[question-1,alternativeCounter]/numUpper_loc  
            middlePerc = numQuestionsAlternativesMiddle_loc[question-1,alternativeCounter]/numMiddle_loc  
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
        rowCounter+=1

     
def write_histogramQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,scoreQuestionsIndicatedSeries_loc,averageScoreQuestions_loc):
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
    possibleScores=numpy.array([-1.0/5.0,0.0,1.0,1.0+1.0/5.0]) #TODO make parameter
    print possibleScores
    columnCounter = 1
    for possibleScore in possibleScores[0:len(possibleScores)-1]:
        sheetC.write(rowCounter,columnCounter,possibleScore,style = easyxf(style_header)) 
        columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"gemiddelde",style = easyxf(style_header))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"type vraag",style = easyxf(style_header)) 
    rowCounter+=1    
    
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style = easyxf(font_bold+border_right_medium) )
        hist,bins = numpy.histogram(scoreQuestionsIndicatedSeries_loc[:,question-1],bins=possibleScores-1.0/6.0)
        columnCounter+=1    
        for n in hist:        
            if (hist[0]>hist[len(hist)-1] or hist[0]+hist[1]>hist[len(hist)-1]+hist[len(hist)-2]): #more confident in wrong answer than confident in correct answer
                sheetC.write(rowCounter,columnCounter,n,style = easyxf(style_specialAttention))
            else:
                sheetC.write(rowCounter,columnCounter,n)
            columnCounter+=1
        if averageScoreQuestions_loc[question-1]<0:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],2),style = easyxf(style_specialAttention)        )
        else:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],2)) 
        columnCounter+=1
        correctPerc= float(hist[len(hist)-1])/numParticipants
        if correctPerc > tresholdDifficultyQuestionZero:
            sheetC.write(rowCounter,columnCounter,"0",style = easyxf(style_specialAttention+align_horizvertcenter) ) 
        elif correctPerc > tresholdDifficultyQuestionOne:
            sheetC.write(rowCounter,columnCounter,"*",style=easyxf(align_horizvertcenter)) 
        elif correctPerc > tresholdDifficultyQuestionTwo: 
            sheetC.write(rowCounter,columnCounter,"**",style=easyxf(align_horizvertcenter))     
        elif correctPerc > tresholdDifficultyQuestionThree: 
            sheetC.write(rowCounter,columnCounter,"***",style=easyxf(align_horizvertcenter))
        else: 
            sheetC.write(rowCounter,columnCounter,"****",style=easyxf(style_specialAttention+align_horizvertcenter)    )
        rowCounter+=1

def write_scoreStudents(outputbook_loc,nameSheet_loc,permutations_loc,numParticipants_loc,deelnemers_loc, numQuestions_loc,numAlternatives_loc,content_loc,content_colNrs_loc,totalScore_loc,scoreQuestionsIndicatedSeries_loc,columnSeries_loc,matrixAnswers):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)

    columnCounter = 0;
    rowCounter = 0;
    
    #deelnemersnummers
    sheetC.write(rowCounter, 0,"ijkID", style=easyxf(style_header+ border_right_medium) ) 
    rowCounter+=1
    for i in xrange(0,len(deelnemers_loc)):
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[i], style=easyxf(style_header + border_right_medium))
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 0;
    #total score for indicated series
    sheetC.write(rowCounter,columnCounter,"totale score",style=easyxf(style_header))
    rowCounter+=1
    for i in xrange(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,totalScore_loc[i])
        rowCounter+=1
    columnCounter+=1;
    

    rowCounter = 0;
    #indicated series
    sheetC.write(rowCounter,columnCounter,"reeks",style=easyxf(style_header)) 
    rowCounter+=1
    for i in xrange(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,columnSeries_loc[i])
        rowCounter+=1
    columnCounter+=1;
    
    #score for different questions
    #beware scores are stored per question without the permutation; 
    #so for the student the scores have to be back-permutated to the order they got
    rowCounter = 0;
    #write heading    
    columnCounterHeader = columnCounter
    for question in xrange(1,numQuestions_loc+1):
        sheetC.write(rowCounter,columnCounterHeader,"score vraag " + str(question),style=easyxf(style_header))
        columnCounterHeader+=1
        
    columnCounterScoreQuestions = columnCounter   
    rowCounter = 1;  
    for participant in xrange(len(totalScore_loc)): # loop over participants
        columnCounter = columnCounterScoreQuestions;
        score = scoreQuestionsIndicatedSeries_loc[participant,:]
        sorted_score = [score[i-1] for i in permutations_loc[int(columnSeries_loc[participant]-1)]]
        for question in xrange(1,numQuestions_loc+1):
            sheetC.write(rowCounter,columnCounter,sorted_score[question-1])
            columnCounter+=1
        rowCounter+=1            
 
    #answer for different questions and alternatives
    for question in xrange(1,numQuestions_loc+1):
            rowCounter=0
            sheetC.write(rowCounter,columnCounter,"antwoord vraag " + str(question),style=easyxf(style_header) ) 
            #print "antwoord vraag " + str(question) + str(alternative)
            answer = matrixAnswers[:,question-1]
            rowCounter+=1
            #print 
            for i in xrange(len(totalScore_loc)):
                sheetC.write(rowCounter,columnCounter,answer[i])
                rowCounter+=1                    
            columnCounter+=1;    
            