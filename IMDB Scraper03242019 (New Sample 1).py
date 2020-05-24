# -*- coding: utf-8 -*-
"""
Created on Sun Feb  3 12:18:56 2019

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

os.chdir('C:/Users/Danielove/Desktop/Study/PhD Courses/19spring/SKB/Movie Data 90-17 (Combine sample 1b & 2)')

def getIMDB(n,Marvel,Seq,url):
    html2 = urllib.request.urlopen(url).read()
    soup2 = BeautifulSoup(html2, "lxml")
    s = None; t = None
    try:
        s = soup2.find("div", {"class":"title_wrapper"}).findNext("h1").next_element.strip()
    except:
        pass
    try:
        t = soup2.find("div", {"class":"title_wrapper"}).findNext("a").get_text().strip()
    except:
        pass

# People Directors Writers and first 4 actors
    urlppl = url + 'fullcredits/?ref_=tt_ov_st_sm'
    htmlppl = urllib.request.urlopen(urlppl).read()
    soupppl = BeautifulSoup(htmlppl, "lxml")
    Director = None; Writer = None; Actor = None
    try:
        Director = ', '.join([i.text.strip() for i in soupppl.find("table",{"class":"simpleTable simpleCreditsTable"}).findAll('a')])
    except:
        pass
    try:
        Writer = ', '.join([i.text.strip() for i in soupppl.findAll("table",{"class":"simpleTable simpleCreditsTable"})[1].findAll('a')])
    except:
        pass
    try:
        Actor = ', '.join([i.findPrevious('a').text.strip() for i in soupppl.find("table",{"class":"cast_list"}).findAll('td',{"class":"ellipsis"})][:4])
    except:
        pass
# Return IMDB rating
    RateIMDB = None
    try:
        RateIMDB = soup2.find("span",{"itemprop": "ratingValue"}).get_text()
    except:
        pass
# Return Movie rate
    Rating = soup2.find("div", {"class":"subtext"}).get_text().split('|')[0].strip()
    IMDBName = None; AlsoKnown = None
    try:
        IMDBName = re.search(r'(.*?)\s\(',soup2.find("div",{"class": "title_wrapper"}).findNext('h1').get_text()).group(1).strip()
    except:
        pass
    try:
        AlsoKnown = soup2.find("h4", text = "Also Known As:").nextSibling.strip()
    except:
        pass
# Return Movie length
    Length = None
    try:
        Length = soup2.find("div", {"class":"subtext"}).get_text().split('|')[1].strip()
    except IndexError:
        pass
# Return Movie types and dummies
    Types = ' '.join(soup2.find("div", {"class":"subtext"}).get_text().split('|')[2].split())
    ActDummy = ComDummy = 0
    if "Action".lower() in Types.lower():
        ActDummy = 1
    if "Comedy".lower() in Types.lower():
        ComDummy = 1
# Return Budget
    Budget = None
    try:
        Budget = re.sub('\W+','',soup2.find("h4", {"class": "inline"},text = "Budget:").next_sibling)
    except AttributeError:
        print(Budget)
        pass
# Sublist of Oscar nomination
    OscarNoNum = OscarAwNum = 0
    if len(soup2.findAll("a", text = "See more awards")) != 0:
        urlN = url + "awards?ref_=tt_awd"
        htmlN = urllib.request.urlopen(urlN).read()
        soupN = BeautifulSoup(htmlN, "lxml")
        for L in soupN.findAll("span", {"class": "award_category"},text = "Oscar"):
            if L.parent.find('b').get_text() == "Winner":
                OscarAwNum = L.parent['rowspan']
            if L.parent.find('b').get_text() == "Nominee":
                OscarNoNum = L.parent['rowspan']
            else:
                continue
# Step3  filter the needed dates;
    url3 = url + "releaseinfo?ref_=tt_ov_inf"
    html3 = urllib.request.urlopen(url3).read()
    soup3 = BeautifulSoup(html3, "lxml")
    RLDList = soup3.find("table", {"class": "ipl-zebra-list ipl-zebra-list--fixed-first release-dates-table-test-only"}).findAll('tr')
    Index = []
    for i in range(len(RLDList)):
        try:
            time.strptime(RLDList[i].find('td', {"class": "release-date-item__date"}).get_text(),"%d %B %Y")
        except:
            print("Wrong time format")
            Index.append(i)
    RLDList = [i for j, i in enumerate(RLDList) if j not in Index]
    US1 = [];USFN = [];USFestivel = [];USPN = [];USPremier = [];USLimited = [];USInternet = []
    GlobalInternet = [];GlobalInternetC = []
    Global1 = [];Global1C = [];Global1F = [];Global1FC = [];Global1FInfo = [];Global1P = [];Global1PC = []
    Global1PInfo = [];Global1L = [];Global1LC = []
    USType = [];US1stDate = [];GlobalType = [];Global1stCountry = [];Global1stDate = []
    for RL in RLDList:
        Global1stDate.append(RL.find('td',{"class": "release-date-item__date"}).get_text())
        Global1stCountry.append(RL.find("a").get_text().strip())
        if RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text() == "":
            GlobalType.append("General")
        elif "festival" in RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text().lower():
            GlobalType.append("Festival")
        elif "premiere" in RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text().lower():
            GlobalType.append("Premiere")
        elif "limited" in RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text().lower():
            GlobalType.append("Limited")
        elif "internet" in RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text().lower():
            GlobalType.append("Internet")
        if "USA" in RL.find("a").get_text():
            US1stDate.append(RL.find('td',{"class": "release-date-item__date"}).get_text())
            if RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text() == "":
                USType.append("General")
            elif "festival" in RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text().lower():
                USType.append("Festival")
            elif "premiere" in RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text().lower():
                USType.append("Premiere")
            elif "limited" in RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text().lower():
                USType.append("Limited")
            elif "internet" in RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text().lower():
                USType.append("Internet")
        
        if "USA" in RL.find("a").get_text() and RL.find('td', {"class": "release-date-item__date"}).findNext('td').\
        get_text() == "":
            US1.append(RL.find('td',{"class": "release-date-item__date"}).get_text())
        if "USA" in RL.find("a").get_text() and "festival" in RL.find('td', {"class": "release-date-item__date"}).\
        findNext('td').get_text().lower():
            USFestivel.append(RL.find('td',{"class": "release-date-item__date"}).get_text())
            USFN.append(re.search("\((.*?)\)",RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text()).group(1))
        if "USA" in RL.find("a").get_text() and "premiere" in RL.find('td', {"class": "release-date-item__date"}).\
        findNext('td').get_text().lower():
            USPremier.append(RL.find('td',{"class": "release-date-item__date"}).get_text())
            USPN.append(re.search("\((.*?)\)",RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text()).group(1))
        if "USA" in RL.find("a").get_text() and "limited" in RL.find('td', {"class": "release-date-item__date"}).\
        findNext('td').get_text().lower():
            USLimited.append(RL.find('td',{"class": "release-date-item__date"}).get_text())
        if "USA" in RL.find("a").get_text() and "internet" in RL.find('td', {"class": "release-date-item__date"}).\
        findNext('td').get_text().lower():
            USInternet.append(RL.find('td',{"class": "release-date-item__date"}).get_text())
        if RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text() == "":
            Global1.append(RL.find('td',{"class": "release-date-item__date"}).get_text())
            Global1C.append(RL.find("a").get_text().strip())
        if "premiere" in RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text().lower():
            Global1P.append(RL.find('td',{"class": "release-date-item__date"}).get_text())
            Global1PC.append(RL.find("a").get_text().strip())
            Global1PInfo.append(re.search("\((.*?)\)",RL.find('td', {"align": "right"}).findNext('td').get_text()).group(1))
        if "festival" in RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text().lower():
            Global1F.append(RL.find('td',{"class": "release-date-item__date"}).get_text())
            Global1FC.append(RL.find("a").get_text().strip())
            Global1FInfo.append(re.search("\((.*?)\)",RL.find('td', {"align": "right"}).findNext('td').get_text()).group(1))
        if "limited" in RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text().lower():
            Global1L.append(RL.find('td',{"class": "release-date-item__date"}).get_text())
            Global1LC.append(RL.find("a").get_text().strip())
        if "internet" in RL.find('td', {"class": "release-date-item__date"}).findNext('td').get_text().lower():
            GlobalInternet.append(RL.find('td',{"class": "release-date-item__date"}).get_text())
            GlobalInternetC.append(RL.find("a").get_text().strip())
    const = url.split('/')[4]
# Return the list
    RT_IMDB = [n,s,t,Marvel,Seq,const,RateIMDB,Rating,Length,Types,ActDummy,ComDummy,Budget,OscarAwNum,OscarNoNum,\
               (USType or [None])[0],(US1stDate or [None])[0],(US1 or [None])[0],(USFestivel or [None])[0],\
               (USFN or [None])[0],(USPremier or [None])[0],(USPN or [None])[0],(USLimited or [None])[0],\
               (USInternet or [None])[0],(GlobalType or [None])[0],(Global1stCountry or [None])[0],\
               (Global1stDate or [None])[0],(Global1C or [None])[0],(Global1 or [None])[0],\
               (Global1FC or [None])[0],(Global1FInfo or [None])[0],(Global1F or [None])[0],(Global1PC or [None])[0],\
               (Global1PInfo or [None])[0],(Global1P or [None])[0],(Global1LC or [None])[0],(Global1L or [None])[0],\
               (GlobalInternetC or [None])[0],(GlobalInternet or [None])[0],url,IMDBName,AlsoKnown,Director,Writer,Actor]
    return(RT_IMDB)

# Run function import movies and dates
FMovieList = pd.read_excel('AllMovies9018.xlsx')
MovieID = FMovieList['MovieID']; Marvel = FMovieList['Marvel Movies']; Sequel = FMovieList['Sequel']; URL = FMovieList['URL']
Movie_IMDB = []
for i in range(0,len(FMovieList)):
    try:
        Movie_IMDB.append(getIMDB(str(MovieID[i]),str(Marvel[i]),str(Sequel[i]),str(URL[i])))
        print(i+1, "/", len(MovieID))
        print(getIMDB(str(MovieID[i]),str(Marvel[i]),str(Sequel[i]),str(URL[i])))
    except:
        continue
    
labels2 = ['MovieID','Movie', 'Year','Marvel Movies','Sequel','MovieConst','RateIMDB','Rating','Length','Types','ActDummy','ComDummy','Budget',\
           'OscarAwNum','OscarNoNum','US 1st Release Type','US 1st Release Date',\
           'USGeneral','USFestival','USFestivalInfo','USPremiere','USPremiereInfo','USLimited','USInternet',\
           'Global 1st Release Type','Global 1st Release Country','Global 1st Release Date','GlobalCountry',\
           'GlobalGeneral','GlobalFestivalCountry','GlobalFestivalInfo','GlobalFestival','GlobalPremiereCountry',\
           'GlobalPremiereInfo','GlobalPremiere','GlobalLimitedCountry','GlobalLimited','GlobalInternetCountry',\
           'GlobalInternet','IMDB URL','Name on IMDB','Other Name','Director','Writer','Actor']
df_IMDB = pd.DataFrame.from_records(Movie_IMDB, columns=labels2)

# Change Date Format

def DateT(t):
    if t == None:
        return None
    else:
        return '-'.join(map(str,time.strptime(t,"%d %B %Y")[0:3]))
    
for i,j in itertools.product([16, 17, 18, 20, 22, 23, 26, 28, 31, 34, 36, 38], range(len(df_IMDB))):
    df_IMDB.iloc[j,i] = DateT(df_IMDB.iloc[j,i] or None)
    
# Save

writer = pd.ExcelWriter('IMDB All Movies (90-18).xlsx')
df_IMDB.to_excel(writer,'Sheet1', index=False)
writer.save()