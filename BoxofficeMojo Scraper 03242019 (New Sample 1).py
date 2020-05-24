# -*- coding: utf-8 -*-
"""
Created on Thu May 17 23:01:08 2018

@author: Chao Wang
"""

import os
from bs4 import BeautifulSoup
import urllib
import re
import pandas as pd
import numpy as np
import itertools
import time
from difflib import SequenceMatcher
from datetime import datetime
import unicodedata

os.chdir('C:/Users/Danielove/Desktop/Study/PhD Courses/19spring/SKB/Movie Data 90-17 (Combine sample 1b & 2)')


# 1st step filter all movies with dates or similar name within searching its name from the dataset

def getMojo(n,s,t,RLD,m,n2):
    url = 'http://www.boxofficemojo.com/search/?q=' + urllib.parse.quote(s)
    html = urllib.request.urlopen(url).read()
    soup = BeautifulSoup(html, "lxml")
    print("Film name: "+ str(s))
    S = re.sub('[^A-Za-z0-9]+', ' ', s).strip()
    M = re.sub('[^A-Za-z0-9]+', ' ', m).strip()
    N2 = re.sub('[^A-Za-z0-9]+', ' ', n2).strip()
    print([S,M,N2])

    MojoName = None
    url2 = None
    pool = soup.findAll('tr', bgcolor=lambda x: x in ['#FFFFFF','#FFFF99','#F4F4FF'])
    Len = len(pool)
    try:
        x = 0
        while x < Len:
            try:
                if datetime.strptime(pool[x].findNext('td').findNext('td').findNext('td').findNext('td').findNext('td').\
                    findNext('td').findNext('td').get_text(), "%m/%d/%Y").strftime('%Y-%#m-%#d') in RLD:
                    MojoName = pool[x].findNext('a').findNext('a').get_text()
                    url2 = 'http://www.boxofficemojo.com' + pool[x].findNext('a')['href'].replace(u'\xa0', u' ')
                    break
                x+=1
            except:
                x+=1
                pass
        if url2 == None:
            x = 0
            while x < Len:            
                if pool[x].findAll('a')[1].get_text() in [s,m,n2,S,M,N2]:
                    MojoName = pool[x].findNext('a').findNext('a').get_text()
                    url2 = 'http://www.boxofficemojo.com' + pool[x].findNext('a')['href'].replace(u'\xa0', u' ')
                    break
                x+=1
            if url2 == None:
                x = 0
                while x < Len:                            
                    if SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),s).ratio() > 0.8 or\
                        SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),str(s+" ("+(str(t)+")"))).ratio() > 0.8 or\
                        SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),S).ratio() > 0.8 or\
                        SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),str(S+" ("+(str(t)+")"))).ratio() > 0.8:
                        MojoName = pool[x].findNext('a').findNext('a').get_text()
                        url2 = 'http://www.boxofficemojo.com' + pool[x].findNext('a')['href'].replace(u'\xa0', u' ')
                        break
                    elif SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),m).ratio() > 0.8 or\
                        SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),str(m+" ("+(str(t)+")"))).ratio() > 0.8 or\
                        SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),M).ratio() > 0.8 or\
                        SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),str(M+" ("+(str(t)+")"))).ratio() > 0.8:
                        MojoName = pool[x].findNext('a').findNext('a').get_text()
                        url2 = 'http://www.boxofficemojo.com' + pool[x].findNext('a')['href'].replace(u'\xa0', u' ')
                        break
                    elif SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),n2).ratio() > 0.8 or\
                        SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),str(n2+" ("+(str(t)+")"))).ratio() > 0.8 or\
                        SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),N2).ratio() > 0.8 or\
                        SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),str(N2+" ("+(str(t)+")"))).ratio() > 0.8:
                        MojoName = pool[x].findNext('a').findNext('a').get_text()
                        url2 = 'http://www.boxofficemojo.com' + pool[x].findNext('a')['href'].replace(u'\xa0', u' ')
                        break
                    x+=1
                if url2 == None:
                    print("Fail")
                    return([])
    except:
        return([])

    html2 = urllib.request.urlopen(url2).read()
    soup2 = BeautifulSoup(html2, "lxml")
    TOTDom = None
    try:
        TOTDom = re.sub("\W+","",soup2.body.find(text='Domestic:').findNext("td").get_text())
    except:
        print(s, " has no Domestic Gross")
        pass
    TOTGlobal = None
    try:
        TOTGlobal = re.sub("\W+","",soup2.body.find(text='Worldwide:').findNext("td").get_text())
    except:
        print(s, " has no Global Gross")
        pass
    WE = None
    try:
        WE = 'http://www.boxofficemojo.com' + soup2.findAll("a", text = "Weekend")[1]['href']
    except:
        print(s, " has no Weekly Gross")
        pass
    WE14 = None; WELast = None; OPWE = None; LENinWs = None
    if WE != None:
        htmlWE = urllib.request.urlopen(WE).read()
        soupWE = BeautifulSoup(htmlWE, "lxml")
        OPWE = re.sub('\W+','',soupWE.findAll("table", {"class": "chart-wide"})[0].findAll("tr")[1].\
                      findAll("td")[2].get_text())
        LENinWs = len(soupWE.findAll('tr', bgcolor=lambda x: x in ['#ffffff','#f4f4ff'])[4:])
        if LENinWs >= 14:
            WE14 = re.sub('\W+','',soupWE.findAll('tr', bgcolor=lambda x: x in ['#ffffff','#f4f4ff'])\
                    [4:][13].findAll("td")[7].get_text())
            WELast = WE14
        else:
            WELast = re.sub('\W+','',soupWE.findAll('tr', bgcolor=lambda x: x in ['#ffffff','#f4f4ff'])\
                    [4:][LENinWs-1].findAll("td")[7].get_text())
    print([n,s,MojoName,t,OPWE,WE14,LENinWs,WELast,TOTDom,TOTGlobal,url2])
    return([n,s,MojoName,t,OPWE,WE14,LENinWs,WELast,TOTDom,TOTGlobal,url2])




# 2nd step filter all movies with dates or similar name within searching its altname from the IMDB.com

def getMojo2(n,s,t,RLD):
    url = 'http://www.boxofficemojo.com/search/?q=' + urllib.parse.quote(s)
    html = urllib.request.urlopen(url).read()
    soup = BeautifulSoup(html, "lxml")
    MojoName = None
    url2 = None
    pool = soup.findAll('tr', bgcolor=lambda x: x in ['#FFFFFF','#FFFF99','#F4F4FF'])
    Len = len(pool)
    try:
        x = 0
        while x < Len:
            try:
                if datetime.strptime(pool[x].findNext('td').findNext('td').findNext('td').findNext('td').findNext('td').\
                    findNext('td').findNext('td').get_text(), "%m/%d/%Y").strftime('%Y-%#m-%#d') in RLD:
                    MojoName = pool[x].findNext('a').findNext('a').get_text()
                    url2 = 'http://www.boxofficemojo.com' + pool[x].findNext('a')['href'].replace(u'\xa0', u' ')
                    break
                x+=1
            except:
                x+=1
                pass
        if url2 == None:
            x = 0
            while x < Len:
                if pool[x].findAll('a')[1].get_text() in [s,re.sub('[^A-Za-z0-9]+', ' ', s).strip()]:
                    MojoName = pool[x].findNext('a').findNext('a').get_text()
                    url2 = 'http://www.boxofficemojo.com' + pool[x].findNext('a')['href'].replace(u'\xa0', u' ')
                    break
                elif SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),s).ratio() > 0.8 or\
                    SequenceMatcher(None, pool[x].findAll('a')[1].get_text(),str(s+" ("+(str(t)+")"))).ratio() > 0.8:
                    MojoName = pool[x].findNext('a').findNext('a').get_text()
                    url2 = 'http://www.boxofficemojo.com' + pool[x].findNext('a')['href'].replace(u'\xa0', u' ')
                    break
                x+=1
            if url2 == None:
                print("Fail")
                return([])
    except:
        return([])
    
    html2 = urllib.request.urlopen(url2).read()
    soup2 = BeautifulSoup(html2, "lxml")
    TOTDom = None
    try:
        TOTDom = re.sub("\W+","",soup2.body.find(text='Domestic:').findNext("td").get_text())
    except:
        print(s, " has no Domestic Gross")
        pass
    TOTGlobal = None
    try:
        TOTGlobal = re.sub("\W+","",soup2.body.find(text='Worldwide:').findNext("td").get_text())
    except:
        print(s, " has no Global Gross")
        pass
    WE = None
    try:
        WE = 'http://www.boxofficemojo.com' + soup2.findAll("a", text = "Weekend")[1]['href']
    except:
        print(s, " has no Weekly Gross")
        pass
    WE14 = None; WELast = None; OPWE = None; LENinWs = None
    if WE != None:
        htmlWE = urllib.request.urlopen(WE).read()
        soupWE = BeautifulSoup(htmlWE, "lxml")
        OPWE = re.sub('\W+','',soupWE.findAll("table", {"class": "chart-wide"})[0].findAll("tr")[1].\
                      findAll("td")[2].get_text())
        LENinWs = len(soupWE.findAll('tr', bgcolor=lambda x: x in ['#ffffff','#f4f4ff'])[4:])
        if LENinWs >= 14:
            WE14 = re.sub('\W+','',soupWE.findAll('tr', bgcolor=lambda x: x in ['#ffffff','#f4f4ff'])\
                    [4:][13].findAll("td")[7].get_text())
            WELast = WE14
        else:
            WELast = re.sub('\W+','',soupWE.findAll('tr', bgcolor=lambda x: x in ['#ffffff','#f4f4ff'])\
                    [4:][LENinWs-1].findAll("td")[7].get_text())
    print([n,s,MojoName,t,OPWE,WE14,LENinWs,WELast,TOTDom,TOTGlobal,url2])
    return([n,s,MojoName,t,OPWE,WE14,LENinWs,WELast,TOTDom,TOTGlobal,url2])


#1.

FMovieList = pd.read_excel('IMDB All Movies (90-18).xlsx')
#FMovieList = FMovieList.iloc[MovieIDList-1]

Movies = FMovieList['Movie'];Year = FMovieList['Year'];MovieID = FMovieList['MovieID'];IMDBMovies = FMovieList['Name on IMDB']
AltName = FMovieList['Other Name']
USRLDate = FMovieList.iloc[:,[14,15,16,18,20,21]]
MojoOUT = None
Movie_Mojo = []
for i in range(0,len(Movies)):
    try:
        MojoOUT = getMojo(str(MovieID[i]),str(Movies[i]),str(Year[i]),list(USRLDate.iloc[i]),str(IMDBMovies[i]),str(AltName[i]))
    except:
        pass
    Movie_Mojo.append(MojoOUT)
    print(i+1, "/", len(Movies))

labels = ['MovieID','Movie','Name on Mojo', 'Year',"BOX1WK","BOX14ORLESS","MOVTHWKS","BOX14WKS","BOXUSLIFE","BOXGLOBALLIFE","Boxoffice URL"]
df_Mojo = pd.DataFrame.from_records(list(filter(None, Movie_Mojo)), columns=labels).drop_duplicates()


# 2.1 Add Movies not scraped and search on IMDB.com

FMovieList2 = FMovieList[-FMovieList['MovieID'].isin(list(df_Mojo["MovieID"]))]
FMovieList2 = pd.DataFrame(FMovieList2.values, columns = list(FMovieList.columns))

Movies2 = FMovieList2['Movie'];Year2 = FMovieList2['Year'];MovieID2 = FMovieList2['MovieID'];IMDBMovies2 = FMovieList2['Name on IMDB']
AltName2 = FMovieList2['Other Name']
USRLDate2 = FMovieList2.iloc[:,[14,15,16,18,20,21]]
MojoOUT2 = None
Movie_Mojo2 = []
#VoidlistMojo2 = []

for i in range(0,len(IMDBMovies2)):
    try:
        MojoOUT2 = getMojo2(str(MovieID2[i]),str(IMDBMovies2[i]),str(Year2[i]),list(USRLDate2.iloc[i]))
    except:
        pass
    Movie_Mojo2.append(MojoOUT2)
    print(i+1, "/", len(Movies2))

df_Mojo2 = pd.DataFrame.from_records(list(filter(None, Movie_Mojo2)), columns=labels).drop_duplicates()


# 2.2 Search with Other Name on IMDB.com

FMovieList3 = FMovieList[-FMovieList['MovieID'].isin(list(df_Mojo["MovieID"]) + list(df_Mojo2["MovieID"]))]
FMovieList3 = pd.DataFrame(FMovieList3.values, columns = list(FMovieList.columns))

Movies3 = FMovieList3['Movie'];Year3 = FMovieList3['Year'];MovieID3 = FMovieList3['MovieID'];IMDBMovies3 = FMovieList3['Name on IMDB']
AltName3 = FMovieList3['Other Name']
USRLDate3 = FMovieList3.iloc[:,[14,15,16,18,20,21]]
MojoOUT3 = None
Movie_Mojo3 = []
#VoidlistMojo2 = []
for i in range(0,len(AltName3)):
    try:
        MojoOUT3 = getMojo2(str(MovieID3[i]),str(AltName3[i]),str(Year3[i]),list(USRLDate3.iloc[i]))
    except:
        pass
    Movie_Mojo3.append(MojoOUT3)
    print(i+1, "/", len(Movies3))

df_Mojo3 = pd.DataFrame.from_records(list(filter(None, Movie_Mojo3)), columns=labels).drop_duplicates()



# Aggregate and export

FMovieList3 = pd.concat([df_Mojo,df_Mojo2,df_Mojo3])
FMovieList3['MovieID'] = FMovieList3['MovieID'].astype(int)
FMovieList3 = pd.DataFrame(FMovieList3.sort_values(by=['MovieID']).values,columns=labels)

NotAvailableData = FMovieList[-FMovieList['MovieID'].isin(FMovieList3["MovieID"])].iloc[:,[0,1,2]]

writer = pd.ExcelWriter('Boxoffice All Movies (90-18).xlsx')
FMovieList3.to_excel(writer,'Boxoffce', index=False)
NotAvailableData.to_excel(writer,'Not Available', index=False)
writer.save()

