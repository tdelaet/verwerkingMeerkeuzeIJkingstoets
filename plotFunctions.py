# -*- coding: utf-8 -*-
"""
Created on Sat Aug 26 20:48:54 2023

@author: u0046457
"""
import matplotlib.pyplot as plt
import numpy
import matplotlib

def plotHistogram(saveNameFig,maxScore,totalScore,percPassed):
    numParticipants=totalScore.size
    font = {'family' : 'normal','size'   : 12}
    fig=plt.figure(figsize=(15, 5))
    ax=fig.add_subplot(111)
    n, bins, patches = plt.hist(totalScore,bins=numpy.arange(0-0.5,maxScore+1,1))
    plt.xlabel("score (max " + str(maxScore)+ ")")
    plt.xlim([0-0.5,maxScore+0.5])
    plt.xticks(numpy.arange(1,maxScore+1))
    plt.ylabel("aantal studenten")       
    plt.text(0.966,0.9, 
              'gemiddelde: ' + str(round(numpy.average(totalScore),2)) + "\n" +
              'mediaan: ' + str(int(numpy.median(totalScore)))  + "\n" +
              'percentage geslaagd: ' + str(int(percPassed)) + "%"  + "\n" +
              'aantal deelnemers: ' + str(numParticipants)
              ,transform=ax.transAxes,
            horizontalalignment='right',
            verticalalignment='top',
            bbox=dict(facecolor='none', edgecolor='black', boxstyle='round,pad=1'),
            fontsize=12)     
    matplotlib.rc('font', **font)       
    plt.savefig(saveNameFig, bbox_inches='tight',dpi=300)



    
def plotPieParticipants(saveNameFig,instellingen,numParticipants_all):   
    ##################PLOTTING##################
    def my_autopct(pct):
        total=sum(numParticipants_all)
        val=int(pct*total/100.0)
        return '{p:.2f}%  ({v:d})'.format(p=pct,v=val)
    if (len(instellingen)!=1):
        # plot the pie diagram of the different locations
        plt.figure()
        labels = instellingen    
        plt.pie(numParticipants_all, labels=labels,
                        autopct=my_autopct, shadow=True, startangle=90)    
        plt.title('Aantal deelnemers', bbox={'facecolor':'0.8', 'pad':5})
        plt.savefig(saveNameFig, bbox_inches='tight',dpi=300)