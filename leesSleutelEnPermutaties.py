# -*- coding: utf-8 -*-
"""
Created on Fri Jun 20 13:38:02 2014

@author: tdelaet
"""


import numpy
import os


def leesPermutaties(jaar_loc,toets_loc,numSeries_loc,texinputFolder_loc):     
    name_basis_loc = jaar_loc + "_" + toets_loc + "_IDreeks"
    questions_loc = []
    for serie in xrange(1,numSeries_loc+1):
        name = name_basis_loc+ str(serie)
        questions_loc.append(numpy.loadtxt(texinputFolder_loc + name+".tex",delimiter='\t',dtype=numpy.str))
    
    # check if all have the same length 
    numQuestions_loc=len(questions_loc[0]) 
    
    for serie in xrange(1,numSeries_loc+1):
        if len(questions_loc[serie-1]) != numQuestions_loc:
            print "ERROR: lijsten met indices van vragen hebben niet dezelfde lengte"
    
    permutations_loc = numpy.zeros(numQuestions_loc* numSeries_loc)
    permutations_loc = permutations_loc.reshape( numSeries_loc,numQuestions_loc)

    for question in xrange(1,numQuestions_loc+1):
        for serie in xrange(1,numSeries_loc+1):
            #find label
            indexQuestion=numpy.where(questions_loc[serie-1]==questions_loc[0][question-1])[0][0]
            permutations_loc[serie-1][indexQuestion] = int(question)
    return permutations_loc
    
def leesSleutel(jaar_loc,toets_loc,texinputFolder_loc):
    name_sleutel = jaar_loc + "_" + toets_loc + "_SLEUTELreeks1"
    sleutel = numpy.loadtxt(texinputFolder_loc + name_sleutel+ ".tex",delimiter='\t',dtype=numpy.str)
    sleutel = sleutel[1:len(sleutel):3]
    return sleutel
    
def leesNamenVragen(jaar_loc,toets_loc,texinputFolder_loc,numQuestions_loc):
    name_basis_loc = jaar_loc + "_" + toets_loc + "_IDreeks1"        
    questionNames = [[] for i in range(int(numQuestions_loc))]
    if not os.path.isfile(texinputFolder_loc + name_basis_loc + ".tex"):
        for question in xrange(0,numQuestions_loc):
            questionNames[question] = 'vraag' + str(int(question+1))
    else:
        questionNames = numpy.loadtxt(texinputFolder_loc + name_basis_loc+".tex",delimiter='\t',dtype=numpy.str)   
    return questionNames  
    
def leesClassificatieVragen(jaar_loc,toets_loc,texinputFolder_loc,numQuestions_loc):
    name_basis_loc = jaar_loc + "_" + toets_loc + "_CLASSreeks1"  
    questionNames = [[] for i in range(int(numQuestions_loc))]
    if not os.path.isfile(texinputFolder_loc + name_basis_loc + ".tex"):
        for question in xrange(0,numQuestions_loc):
            questionNames[question] = '?'
    else:        
        questionNames = numpy.loadtxt(texinputFolder_loc + name_basis_loc+".tex",delimiter='\t',dtype=numpy.str)   
    return questionNames  
    
    
def leesCategorieVragen(jaar_loc,toets_loc,texinputFolder_loc,numQuestions_loc):    
    name_basis_loc = jaar_loc + "_" + toets_loc + "_CATEGORIENreeks1"
    questionCategorie = [[] for i in range(int(numQuestions_loc))]
    if not os.path.isfile(texinputFolder_loc + name_basis_loc + ".tex"):
        for question in xrange(0,numQuestions_loc):
            questionCategorie[question] = 'vraag' + str(int(question+1))
    else:
        questionCategorie = numpy.loadtxt(texinputFolder_loc + name_basis_loc+".tex",delimiter='\t',dtype=numpy.str)
    return questionCategorie
    