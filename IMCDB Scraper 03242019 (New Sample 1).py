# -*- coding: utf-8 -*-
"""
Created on Mon Feb  4 00:09:18 2019

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

#n,s,t,IMDB_url = ["21","Be Cool","2005","tt0377471"]
Voidlist = dict([(key, []) for key in ["A","B"]])
def MatchIMCDB(n,s,t,IMDB_url,RLDate):
    try:
        url = 'http://www.imcdb.org/movie_' + IMDB_url[2:] + '.html'
        IMCDBhtml = urllib.request.urlopen(url).read()
        IMCDBsoup = BeautifulSoup(IMCDBhtml, "lxml")
 # No Record of the Movie
    except:
        Voidlist['A'].append([n,s,t])
        return([])
    Model_list = IMCDBsoup.find_all('span', {"class": 'Stars'})
    print("There are ", len(Model_list), " stars in movie: " + s)
    Model = []; i = 0; 
    for listitem in Model_list:
        try:
            if listitem.find('img')['alt'] == "[*]":
                try:
                    year = str(re.search("\d{4}",str(listitem.findParent("div").findNext("a").\
                                         next_sibling)).group(0))
                except:
                    year = None
                urlcar = 'http://www.imcdb.org/' + listitem.findParent("div").findNext("a")['href']
                htmlcar = urllib.request.urlopen(urlcar).read()
                soupcar = BeautifulSoup(htmlcar, "lxml")
                try:
                    classes = re.search(r"\: (.*?) \—",soupcar.find('div', {"id":"VehicleDetails"}).findNext('p').\
                                        get_text()).group(1).replace("\t", " ")
                except:
                    classes = None
                try:
                    origin = soupcar.find('div', {"id":"VehicleDetails"}).findNext('img',{"class":"CarFlag"})['title']
                except:
                    origin = None
                try:
                    makefor = soupcar.find('div', {"id":"VehicleDetails"}).findNext('img',{"class":"CarFlag"}).\
                    findNext('img',{"class":"CarFlag"})['title']
                except:
                    makefor = None
                add = listitem.findParent("div").findNext("a").findNext("a").get_text()
                model = listitem.findParent("div").findNext("a").findNext("a").findNext("a").get_text()
                nstar = len(listitem.find_all('img'))
                print(add)
                Model.append([year,add,model,classes,origin,makefor,nstar,len(Model_list),n,s,t,RLDate,url])
                i += 1
        except:
            pass
    if i == 0:
# No Stared Cars in the Movie
        Voidlist['B'].append([n,s,t,url])
        return([])
    else:
        return(Model)


FMovieList = pd.read_excel('IMDB All Movies (90-18).xlsx')
URLconst = FMovieList['MovieConst']; Movies = FMovieList['Movie'];Year = FMovieList['Year']
MovieID = FMovieList['MovieID']
USRLDate = FMovieList['US 1st Release Date']
AutoADD = None
Auto_Movie = []
for i in range(0,len(Movies)):
    try:
        AutoADD = MatchIMCDB(str(MovieID[i]),str(Movies[i]),str(Year[i]),str(URLconst[i]),str(USRLDate[i]))
        Auto_Movie.append(AutoADD)
        print(i+1, "/", len(Movies))
    except:
        pass


labelCarMovies = ['Vehicle Year','Make','Model','Classes','Origin','Make for','Number of Stars','Number of Cars','MovieID','Movie Name',\
                  'Movie Year','US Release Date','IMCDB URL']
df_CarMovie = pd.DataFrame.from_records(list(itertools.chain.from_iterable(Auto_Movie)), columns=labelCarMovies)
labelA = ['MovieID','Movie','Year']
ListA = pd.DataFrame.from_records(list(Voidlist['A']), columns=labelA)
labelB = ['MovieID','Movie','Year','IMCDB URL']
ListB = pd.DataFrame.from_records(list(Voidlist['B']), columns=labelB)

writer = pd.ExcelWriter('IMCDB All Movies (90-18).xlsx')
df_CarMovie.to_excel(writer,'IMCDB Data', index=False)
ListA.to_excel(writer,'No Result Searching IMCDB', index=False)
ListB.to_excel(writer,'No Stared Cars of the Movie', index=False)
writer.save()
