# -*- coding: utf-8 -*-
"""
Created on Wed May  9 11:12:31 2018

@author: Chao Wang
"""

# 2.3 Blu-ray Varlist(RELDATE (Global),RELEASEDATE,MOVBUDG,CRITRAT,MOVGENRES,DUMACTION,DUMCOM,
# Use proxies

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
import requests


os.chdir('C:/Users/Danielove/Desktop/Study/PhD Courses/19spring/SKB/Movie Data 90-17 (Combine sample 1b & 2)')

#n,url,s,t,proxy = ['5','https://www.imdb.com/title/tt0117998/?ref_=adv_li_tt'[:-15],'Twister','1996','176.215.237.190:55271']
def getBlu(n,url,s,t):
#    Parse the main page searching the movie's Blue-ray products
    if pd.isnull(s):
        return([n,s,t,None,None,None,None,url])
    s = re.sub('[^a-zA-Z0-9 \n\.]', '', s).strip()
    url1 = 'http://www.blu-ray.com/movies/search.php?action=search&keyword='+urllib.parse.quote(s)+'&yearfrom='+str(int(t)-1)+\
    '&yearto='+str(int(t)+1)+'&sortby=releasetimestampdesc'
    html = requests.get(url1).content
    soup = BeautifulSoup(html, "lxml")
    print("Film name: "+ str(s))
    BluRayRL = None
    BluRayLink = None
#    Get all the Blue-ray products
    try:
        if len(soup.findAll("td", {"width":"76%"})) == 0:
            Bluelist = [str(url1) + '&page=' + "0"]
        else:
            Bluelist = [str(url1) + '&page=' + str(i) for i in range(len(soup.find("td", {"width":"76%"}).find_all('a'))+1)]
        page = 0
#        Find the 1st Blue-ray product matches the movie's record on www.IMDB.com
        while page < len(Bluelist):
            print(str(url1)+'&page='+str(page))
            html2 = requests.get(Bluelist[page]).content
            soup2 = BeautifulSoup(html2, "lxml")
            Cands = soup2.findAll("table", {"width":"113", "border":"0", "cellpadding":"0", "cellspacing":"0", "style":\
                                            "display: inline-block"})
            if len(Cands) != 0:
                NCands = 0
                while NCands < len(Cands):
                    try:
                        urlcand = Cands[NCands].find('a')['href']
                        print(NCands,"/",len(Cands)," ",urlcand)
                        htmlc = requests.get(urlcand).content
                        soupc = BeautifulSoup(htmlc, "lxml")
                        if soupc.find('a',{"id":"imdb_icon"})['href'] == url:
                            if "Blu" in soupc.find('td',{"width":"228px","style":"font-size: 12px"}).get_text():
                                BluRayRL = '-'.join(map(str,time.strptime(soupc.find('a',{\
                                "style":"text-decoration: none; color: #666666"}).get_text(),"%b %d, %Y")[0:3]))
                                BluRayLink = urlcand
                                break
                        if BluRayRL != None:
                            break
                        NCands += 1
                    except:
                        NCands += 1
                        pass
                    if NCands >= len(Cands):
                        break
            if len(Cands) == 0:
                break
            page += 1
            if page >= len(Bluelist):
                break
    except:
        pass
    return([n,s,t,BluRayRL,BluRayLink,None,None,url])

def getDVD(n,url,s,t):
#    Parse the main page searching the movie's Blue-ray products
    if pd.isnull(s):
        return([n,s,t,None,None,None,None,url])
    s = re.sub('[^a-zA-Z0-9 \n\.]', '', s).strip()
    url1_1 = 'http://www.blu-ray.com/dvd/search.php?action=search&keyword='+urllib.parse.quote(s)+'&yearfrom='+str(int(t)-1)+\
    '&yearto='+str(int(t)+1)+'&sortby=releasetimestampdesc'
    html1_1 = requests.get(url1_1).content
    soup1_1 = BeautifulSoup(html1_1, "lxml")
    print("Film name: "+ str(s))
    DVDRL = None
    DVDLink = None
#    Get all the DVD products
    try:
        if len(soup1_1.findAll("td", {"width":"76%"})) == 0:
            Bluelist_1 = [str(url1_1) + '&page=' + "0"]
        else:
            Bluelist_1 = [str(url1_1) + '&page=' + str(i) for i in range(len(soup1_1.find("td", {"width":"76%"}).find_all('a')))]
        page_1 = 0
#        Find the 1st DVD product matches the movie's record on www.IMDB.com
        while page_1 < len(Bluelist_1):
            print(str(url1_1)+'&page='+str(page_1))
            html2_1 = requests.get(Bluelist_1[page_1]).content
            soup2_1 = BeautifulSoup(html2_1, "lxml")
            Cands_1 = soup2_1.findAll("table", {"width":"113", "border":"0", "cellpadding":"0", "cellspacing":"0", "style":\
                                            "display: inline-block"})
            if len(Cands_1) != 0:
                NCands2 = 0
                while NCands2 < len(Cands_1):
                    try:
                        urlcand_1 = Cands_1[NCands2].find('a')['href']
                        print(NCands2,"/",len(Cands_1)," ",urlcand_1)
                        htmlc_1 = requests.get(urlcand_1).content
                        soupc_1 = BeautifulSoup(htmlc_1, "lxml")
                        if soupc_1.find('a',{"id":"imdb_icon"})['href'] == url:
                            print(urlcand_1)
                            if "DVD" in soupc_1.find('td',{"width":"228px","style":"font-size: 12px"}).get_text():
                                DVDRL = '-'.join(map(str,time.strptime(soupc_1.find('a',{\
                                "style":"text-decoration: none; color: #666666"}).get_text(),"%b %d, %Y")[0:3]))
                                DVDLink = urlcand_1
                                break
                        if DVDRL != None:
                            break
                        NCands2 += 1
                    except:
                        NCands2 += 1
                        pass
                    if NCands2 >= len(Cands_1):
                        break
            if len(Cands_1) == 0:
                break
            page_1 += 1
            if page_1 >= len(Bluelist_1):
                break
    except:
        pass
    return([n,s,t,None,None,DVDRL,DVDLink,url])


FMovieList = pd.read_excel('IMDB All Movies (90-18).xlsx')
labels = ['MovieID','Movie','Year','Blue-Ray Release','Blue-Ray Link','DVD Release','DVD Link','IMDB URL']
#
#BlurayToGo = reader[pd.isnull(reader['Blue-Ray Release'])]
#DVDToGo = reader[pd.isnull(reader['DVD Release'])]
# BlurayToGo

# Run function import movies and dates
#Read = pd.read_excel('C:/Users/Danielove/Desktop/Study/PhD Courses/19spring/SKB/Movie Data 94-04/IMDB9004.xlsx')
#labels = list(Read.columns)

# Scrape BlueRay RLD

Movie_Blu = []
Next = []

# We have to try searching Original movie names/Names of its title on IMDB/Other names of the movie

i = 0
while i < len(FMovieList):
    try:
        add = getBlu(str(FMovieList['MovieID'][i]),str(FMovieList['IMDB URL'][i]),str(FMovieList['Movie'][i]),str(FMovieList['Year'][i]))
        if pd.isnull(add[4]):
            add = getBlu(str(FMovieList['MovieID'][i]),str(FMovieList['IMDB URL'][i]),str(FMovieList['Name on IMDB'][i]),str(FMovieList['Year'][i]))
            if pd.isnull(add[4]) and pd.notna(str(FMovieList['Other Name'][i])):
                add = getBlu(str(FMovieList['MovieID'][i]),str(FMovieList['IMDB URL'][i]),str(FMovieList['Other Name'][i]),str(FMovieList['Year'][i]))
                if pd.isnull(add[4]):
                    Next.append(add)
        if pd.notna(add[4]):
            Movie_Blu.append(add)
        print(i+1, "/", len(FMovieList))
        i += 1
    except:
        pass

BlueList = pd.DataFrame.from_records(Movie_Blu, columns=labels)
NextBlue = pd.DataFrame.from_records(Next, columns=labels)

#for i in range(len(BlueList['MovieID'])):
##    print(int(pd.to_numeric(i, downcast='signed')))
#    transfer = int(pd.to_numeric(BlueList['MovieID'][i], downcast='signed'))
##    print(transfer)
#    print(reader.iloc[transfer-1,3:5])
#    reader.iloc[transfer-1,3:5] = BlueList.iloc[i,3:5]
    

# Scrape DVD RLD

Movie_DVD = []
NextD = []

# We have to try searching Original movie names/Names of its title on IMDB/Other names of the movie

i = 0
while i < len(FMovieList):
    try:
        add = getDVD(str(FMovieList['MovieID'][i]),str(FMovieList['IMDB URL'][i]),str(FMovieList['Movie'][i]),str(FMovieList['Year'][i]))
        if pd.isnull(add[6]):
            add = getDVD(str(FMovieList['MovieID'][i]),str(FMovieList['IMDB URL'][i]),str(FMovieList['Name on IMDB'][i]),str(FMovieList['Year'][i]))
            if pd.isnull(add[6]) and pd.notna(str(FMovieList['Other Name'][i])):
                add = getDVD(str(FMovieList['MovieID'][i]),str(FMovieList['IMDB URL'][i]),str(FMovieList['Other Name'][i]),str(FMovieList['Year'][i]))
                if pd.isnull(add[6]):
                    NextD.append(add)
        if pd.notna(add[6]):
            Movie_DVD.append(add)
        print(i+1, "/", len(FMovieList))
        i += 1
    except:
        pass


DVDList = pd.DataFrame.from_records(Movie_DVD, columns=labels)
NextDVD = pd.DataFrame.from_records(NextD, columns=labels)

  
# Bind: BlueList,  NextBlue; DVDList, NextDVD

BlueRays = pd.concat([BlueList,NextBlue])
pd.to_numeric(BlueRays.MovieID, errors='coerce')
DVDs = pd.concat([DVDList,NextDVD])
pd.to_numeric(DVDs.MovieID, errors='coerce')

# Merge

AllData = pd.merge(BlueRays[['MovieID', 'Movie', 'Year', 'Blue-Ray Release', 'Blue-Ray Link', 'IMDB URL']],\
         DVDs[['MovieID','DVD Release','DVD Link']], on='MovieID')
AllData = AllData[['MovieID', 'Movie', 'Year', 'Blue-Ray Release', 'Blue-Ray Link', 'DVD Release', 'DVD Link',\
       'IMDB URL']]
pd.DataFrame(AllData.sort_values(by=['MovieID']).values, columns=labels)

# Write a copy of scraped Blueray & DVD products and move on next trail of names

AllData.MovieID = pd.to_numeric(AllData.MovieID, errors='coerce')
AllData = AllData.sort_values(by=['MovieID'])


writer = pd.ExcelWriter('Blueray All Movies (90-18).xlsx')
AllData.to_excel(writer,str("BlueRay and DVD"), index=False)
writer.save()

