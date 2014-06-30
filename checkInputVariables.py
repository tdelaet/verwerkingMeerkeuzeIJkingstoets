# -*- coding: utf-8 -*-
"""
Created on Wed May 21 14:37:18 2014

@author: tdelaet
"""

from xlrd import open_workbook, biffh


def checkInputVariables(nameFile_loc,nameSheet_loc,numQuestions_loc,numAlternatives_loc,numSeries_loc,correctAnswers_loc,permutations_loc,locations_loc):
    return (
    checkFileAndSheet(nameFile_loc,nameSheet_loc,locations_loc) &
    checkCorrectAnswers(numQuestions_loc, numAlternatives_loc, correctAnswers_loc) & 
    checkPermutations(numSeries_loc,numQuestions_loc, permutations_loc)
    )
            
def checkFileAndSheet(nameFile_loc,nameSheet_loc,locations_loc):
    for location in locations_loc:
        try:
            book = open_workbook(nameFile_loc+"_"+location+ ".xlsx" )
            book.sheet_by_name(nameSheet_loc)
        except IOError:
            print "the selected file " + nameFile_loc +  " can not be opened as a workbook"
            return False
        except biffh.XLRDError:
            print "the selected sheet " + nameSheet_loc +  " can not be opened"
            return False
    return True;    
        
            
def checkCorrectAnswers(numQuestions_loc, numAlternatives_loc, correctAnswers_loc):
    if (len(correctAnswers_loc) == numQuestions_loc):
        if not(set(correctAnswers_loc).issubset(set(map(chr, range(65,65+numAlternatives_loc))))): #correct answers does not only contain A,B,C, .. up to number of alternatives:
            print "ERROR: The list of correct answers " + str(correctAnswers_loc) +  " does not only contain " + str(map(chr, range(65,65+numAlternatives_loc)))
    else:
        print "ERROR: The number of indicated questions " + str(numQuestions_loc) +  " is not equal to number of correct answers listed " + str(correctAnswers_loc)
        return False
    return True             
        
def checkPermutations(numSeries_loc,numQuestions_loc, permutations_loc):             
    # check that for all the permutations all questions are present  
    if (len(permutations_loc) == numSeries_loc):
        # check if all questions are present
        for permutationNumber_loc in xrange(1,numSeries_loc+1):
            if (set(xrange(1,numQuestions_loc+1)) != set(permutations_loc[permutationNumber_loc-1])):
                print "ERROR: Not all " + str(numQuestions_loc) +  " questions are present in permutation " + str(permutationNumber_loc) + ": " + str(permutations_loc[permutationNumber_loc-1])
                return False
    else:
        print "ERROR: The number of indicated series " + str(numSeries_loc) +  " is not equal to the number of permutations listed in the permutation list " + str(permutations_loc)
        return False
    return True     

