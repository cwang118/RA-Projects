
# Sys.setlocale("LC_TIME", "English")
library(openxlsx)
library(timeDate)
library(bizdays)
library(dplyr)
library(stringr)

setwd("C:/Users/Danielove/Desktop/Study/PhD Courses/19spring/SKB/Movie Data 90-18/Scraped Data")
getwd()


# 1. IMDB Data

{
  df_IMDB = read.xlsx("IMDB All Movies (90-18).xlsx",1)
  df_IMDB = df_IMDB[-6]
  names(df_IMDB) = gsub("\\.", " ", names(df_IMDB))
  names(df_IMDB) = gsub("([a-z])([A-Z])", "\\1 \\2", names(df_IMDB))
  names(df_IMDB) = gsub("(US)([A-Z])", "\\1 \\2", names(df_IMDB))
  names(df_IMDB)[c(1,27)] = c("MovieID","Global General Country")
  Trim = function(x){gsub("[^[:alnum:][:blank:]+?&/\\-]", "", x)}
  df_IMDB[,c(23,25,27,30,33,35)+2] = sapply(df_IMDB[,c(23,25,27,30,33,35)+2],Trim)
  
  # Transfer time length from hhmm to min
  
  df_IMDB$Budget = gsub("\\D", "", df_IMDB$Budget)
  HMtoM = function(x) # Transfer hour to min
  {
    X1 = ifelse(str_count(x,"h") == 1, as.numeric(str_match(x, "(.*?)[h*]")[2])*60, 0)
    X2 = ifelse(str_count(x,"min") == 1, as.numeric(str_match(x, "(\\d*)min")[2]), 0)
    X = X1+X2
    return(X)
  }
  df_IMDB$Length = sapply(df_IMDB$Length,HMtoM)
  
  # Get list of Valid Movies and Remove Duplicated Movies
  
  ValidMovieID = df_IMDB[-which(duplicated(df_IMDB$`IMDB URL`)),1]
  
  df_IMDB = df_IMDB[-which(duplicated(df_IMDB$`IMDB URL`)),]
  write.xlsx(df_IMDB,"No Dup Movie 90-18.xlsx")
  
  
}


# 2. IMCDB Data

{
  VehicleData = read.xlsx("IMCDB All Movies (90-18).xlsx",1)
  
  # Remove the dup movies
  VehicleData = VehicleData[which(VehicleData$MovieID %in% ValidMovieID),]
  
  Merged1 = merge(VehicleData[,c(1:9,13)], df_IMDB, by = "MovieID", sort = TRUE)
  names(Merged1) = gsub("\\.", " ", names(Merged1));names(Merged1) = gsub("Festivel", "Festival", names(Merged1))
  Vehicles = Merged1[,c(1,11,12,2:10,24:47)]
  Vehicles = Vehicles[order(as.numeric(Vehicles$MovieID)),]
  
  # Look for the candidate makes
  
  Namelist = c("TATA","Land Rover","Land-Rover","Jaguar","FORD","Lincoln","Cadillac", "Chevrolet", "Buick","GM","HONDA",
               "Acura","Nissan","Infiniti","Toyota","Lexus","Fiat","Daimler","Benz","Chrysler","Dodge","Jeep","Ram",
               "Pontiac")
  TheList = list()
  for (i in 1:length(Namelist))
  {
    TheList[[i]] = names(table(Vehicles$Make))[which(grepl(Namelist[i], names(table(Vehicles$Make)), ignore.case = T))]
  }
  Candidates = unique(unlist(TheList))
  
  # Select from all candidates
  # [1] "Tata"               "Land-Rover"         "Land-Rover Santana" "Jaguar"             "Beauford"          
  # [6] "Bedford"            "Benford"            "Chang'an-Ford"      "Crayford"           "Ford"              
  # [11] "Ford Shelby"        "Fordson"            "Iveco-Ford"         "JMC-Ford"           "Matford"           
  # [16] "Pope-Hartford"      "Radford"            "Lincoln"            "Cadillac"           "Cadillac-Gage"     
  # [21] "Chevrolet"          "Buick"              "McLaughlin-Buick"   "GM"                 "GM Allison"        
  # [26] "GMC"                "GMDD"               "WhiteGMC"           "Dongfeng-Honda"     "Hero Honda"        
  # [31] "Honda"              "Kinetic Honda"      "Acura"              "Dongfeng-Nissan"    "Nissan"            
  # [36] "Nissan Diesel"      "Infiniti"           "FAW-Toyota"         "Guangzhou-Toyota"   "Toyota"            
  # [41] "Lexus"              "Fiat"               "Fiatagri"           "Fiatallis"          "Nanjing-Fiat"      
  # [46] "NSU-Fiat"           "Polski Fiat"        "Austro-Daimler"     "Daimler"            "Daimler-Benz"      
  # [51] "Mercedes-Benz"      "Chrysler"           "Dodge"              "Dodge Brothers"     "Jeep"              
  # [56] "Kaiser Jeep"        "Berliet-Tramagal"   "Framo"              "Holiday Rambler"    "Paramount"         
  # [61] "Ram"                "RAM"                "Rambler"            "Rammax"             "Pontiac"  
  Candidates2 = Candidates[c(1,2,4,10,11,12,18:28,31,33,35:37,40:43,49:56,61:63,65)]
  
  # Final list
  # [1] "Tata"             "Land-Rover"       "Jaguar"           "Ford"             "Ford Shelby"      "Fordson"         
  # [7] "Lincoln"          "Cadillac"         "Cadillac-Gage"    "Chevrolet"        "Buick"            "McLaughlin-Buick"
  # [13] "GM"               "GM Allison"       "GMC"              "GMDD"             "WhiteGMC"         "Honda"           
  # [19] "Acura"            "Nissan"           "Nissan Diesel"    "Infiniti"         "Toyota"           "Lexus"           
  # [25] "Fiat"             "Fiatagri"         "Daimler"          "Daimler-Benz"     "Mercedes-Benz"    "Chrysler"        
  # [31] "Dodge"            "Dodge Brothers"   "Jeep"             "Kaiser Jeep"      "Ram"              "RAM"             
  # [37] "Rambler"          "Pontiac"      
  
  # # List of Qualified Makes
  # df_Candidates = as.data.frame(matrix(NA,length(Candidates),2))
  # df_Candidates[,1] = Candidates
  # df_Candidates[,2] = 0; df_Candidates[c(1,2,4,10,11,12,18:28,31,33,35:37,40:43,49:56,61:63,65),2] = 1
  # names(df_Candidates) = c("Made Name","Qualified Made")
  # df_Candidates[which(df_Candidates[,2]==1),]
  
  matchcar = function(Importdata)
  {
    TATA = rep(NA,6)
    FORD = rep(NA,6)
    GM = rep(NA,6)
    HONDA = rep(NA,6)
    Nissan = rep(NA,6)
    Toyota = rep(NA,6)
    Fiat = rep(NA,6)
    Benz = rep(NA,6)
    
    Car = Importdata[,4]; Release = Importdata[1,ncol(Importdata)];NStar = Importdata[,5]
    TATA[1:2]=FORD[1:2]=GM[1:2]=HONDA[1:2]=Nissan[1:2]=Toyota[1:2]=Fiat[1:2]=Benz[1:2] = 0
    if (Release < "2007-08-17")
    {
      TATAlist = c("Tata")
      if (any(TATAlist %in% Car)) {TATA[1] = 1;TATA[2] = length(which(Car %in% TATAlist))
      TATA[3:6] = as.numeric(summary(NStar[which(Car %in% TATAlist)]))[c(1,4,3,6)]}
      
      FORDlist = c("Ford","Ford Shelby","Fordson","Jaguar","Lincoln","Land-Rover")
      if (any(FORDlist %in% Car))  {FORD[1] = 1;FORD[2] = length(which(Car %in% FORDlist))
      FORD[3:6] = as.numeric(summary(NStar[which(Car %in% FORDlist)]))[c(1,4,3,6)]}
      
      GMlist = c("Pontiac","Cadillac","Cadillac-Gage","Chevrolet","Buick","McLaughlin-Buick","GM","GMC","WhiteGMC","GMDD","GM Allison")
      if (any(GMlist %in% Car)) {GM[1] = 1;GM[2] = length(which(Car %in% GMlist))
      GM[3:6] = as.numeric(summary(NStar[which(Car %in% GMlist)]))[c(1,4,3,6)]}
      
      HONDAlist = c("Honda","Acura")
      if (any(HONDAlist %in% Car)) {HONDA[1] = 1;HONDA[2] = length(which(Car %in% HONDAlist))
      HONDA[3:6] = as.numeric(summary(NStar[which(Car %in% HONDAlist)]))[c(1,4,3,6)]}
      
      Nissanlist = c("Nissan","Nissan Diesel","Infiniti")
      if (any(Nissanlist %in% Car)) {Nissan[1] = 1;Nissan[2] = length(which(Car %in% Nissanlist))
      Nissan[3:6] = as.numeric(summary(NStar[which(Car %in% Nissanlist)]))[c(1,4,3,6)]}
      
      Toyotalist = c("Toyota","Lexus")
      if (any(Toyotalist %in% Car)) {Toyota[1] = 1;Toyota[2] = length(which(Car %in% Toyotalist))
      Toyota[3:6] = as.numeric(summary(NStar[which(Car %in% Toyotalist)]))[c(1,4,3,6)]}
      
      Fiatlist = c("Fiat","Fiatagri")
      if (any(Fiatlist %in% Car)) {Fiat[1] = 1;Fiat[2] = length(which(Car %in% Fiatlist))
      Fiat[3:6] = as.numeric(summary(NStar[which(Car %in% Fiatlist)]))[c(1,4,3,6)]}
      
      Benzlist = c("Daimler","Daimler-Benz","Mercedes-Benz")
      if (any(Benzlist %in% Car)) {Benz[1] = 1;Benz[2] = length(which(Car %in% Benzlist))
      Benz[3:6] = as.numeric(summary(NStar[which(Car %in% Benzlist)]))[c(1,4,3,6)]}
    }
    if (Release >= "2007-08-17" & Release < "2007-09-24") # Add Chrysler for Benz
    {
      TATAlist = c("Tata")
      if (any(TATAlist %in% Car)) {TATA[1] = 1;TATA[2] = length(which(Car %in% TATAlist))
      TATA[3:6] = as.numeric(summary(NStar[which(Car %in% TATAlist)]))[c(1,4,3,6)]}
      
      FORDlist = c("Ford","Ford Shelby","Fordson","Jaguar","Lincoln","Land-Rover")
      if (any(FORDlist %in% Car))  {FORD[1] = 1;FORD[2] = length(which(Car %in% FORDlist))
      FORD[3:6] = as.numeric(summary(NStar[which(Car %in% FORDlist)]))[c(1,4,3,6)]}
      
      GMlist = c("Pontiac","Cadillac","Cadillac-Gage","Chevrolet","Buick","McLaughlin-Buick","GM","GMC","WhiteGMC","GMDD","GM Allison")
      if (any(GMlist %in% Car)) {GM[1] = 1;GM[2] = length(which(Car %in% GMlist))
      GM[3:6] = as.numeric(summary(NStar[which(Car %in% GMlist)]))[c(1,4,3,6)]}
      
      HONDAlist = c("Honda","Acura")
      if (any(HONDAlist %in% Car)) {HONDA[1] = 1;HONDA[2] = length(which(Car %in% HONDAlist))
      HONDA[3:6] = as.numeric(summary(NStar[which(Car %in% HONDAlist)]))[c(1,4,3,6)]}
      
      Nissanlist = c("Nissan","Nissan Diesel","Infiniti")
      if (any(Nissanlist %in% Car)) {Nissan[1] = 1;Nissan[2] = length(which(Car %in% Nissanlist))
      Nissan[3:6] = as.numeric(summary(NStar[which(Car %in% Nissanlist)]))[c(1,4,3,6)]}
      
      Toyotalist = c("Toyota","Lexus")
      if (any(Toyotalist %in% Car)) {Toyota[1] = 1;Toyota[2] = length(which(Car %in% Toyotalist))
      Toyota[3:6] = as.numeric(summary(NStar[which(Car %in% Toyotalist)]))[c(1,4,3,6)]}
      
      Fiatlist = c("Fiat","Fiatagri")
      if (any(Fiatlist %in% Car)) {Fiat[1] = 1;Fiat[2] = length(which(Car %in% Fiatlist))
      Fiat[3:6] = as.numeric(summary(NStar[which(Car %in% Fiatlist)]))[c(1,4,3,6)]}
      
      Benzlist = c("Daimler","Daimler-Benz","Mercedes-Benz","Chrysler","Dodge","Dodge Brothers","Jeep","Kaiser Jeep","Ram","RAM","Rambler")
      if (any(Benzlist %in% Car)) {Benz[1] = 1;Benz[2] = length(which(Car %in% Benzlist))
      Benz[3:6] = as.numeric(summary(NStar[which(Car %in% Benzlist)]))[c(1,4,3,6)]}
    }
    if (Release >= "2007-09-24" & Release < "2008-04-03") # No Fiat
    {
      TATAlist = c("Tata")
      if (any(TATAlist %in% Car)) {TATA[1] = 1;TATA[2] = length(which(Car %in% TATAlist))
      TATA[3:6] = as.numeric(summary(NStar[which(Car %in% TATAlist)]))[c(1,4,3,6)]}
      
      FORDlist = c("Ford","Ford Shelby","Fordson","Jaguar","Lincoln","Land-Rover")
      if (any(FORDlist %in% Car))  {FORD[1] = 1;FORD[2] = length(which(Car %in% FORDlist))
      FORD[3:6] = as.numeric(summary(NStar[which(Car %in% FORDlist)]))[c(1,4,3,6)]}
      
      GMlist = c("Pontiac","Cadillac","Cadillac-Gage","Chevrolet","Buick","McLaughlin-Buick","GM","GMC","WhiteGMC","GMDD","GM Allison")
      if (any(GMlist %in% Car)) {GM[1] = 1;GM[2] = length(which(Car %in% GMlist))
      GM[3:6] = as.numeric(summary(NStar[which(Car %in% GMlist)]))[c(1,4,3,6)]}
      
      HONDAlist = c("Honda","Acura")
      if (any(HONDAlist %in% Car)) {HONDA[1] = 1;HONDA[2] = length(which(Car %in% HONDAlist))
      HONDA[3:6] = as.numeric(summary(NStar[which(Car %in% HONDAlist)]))[c(1,4,3,6)]}
      
      Nissanlist = c("Nissan","Nissan Diesel","Infiniti")
      if (any(Nissanlist %in% Car)) {Nissan[1] = 1;Nissan[2] = length(which(Car %in% Nissanlist))
      Nissan[3:6] = as.numeric(summary(NStar[which(Car %in% Nissanlist)]))[c(1,4,3,6)]}
      
      Toyotalist = c("Toyota","Lexus")
      if (any(Toyotalist %in% Car)) {Toyota[1] = 1;Toyota[2] = length(which(Car %in% Toyotalist))
      Toyota[3:6] = as.numeric(summary(NStar[which(Car %in% Toyotalist)]))[c(1,4,3,6)]}
      
      # Fiatlist = c("Fiat","Fiatagri")
      # if (any(Fiatlist %in% Car)) {Fiat[1] = 1;Fiat[2] = length(which(Car %in% Fiatlist))
      # Fiat[3:6] = as.numeric(summary(NStar[which(Car %in% Fiatlist)]))[c(1,4,3,6)]}
      
      Benzlist = c("Daimler","Daimler-Benz","Mercedes-Benz","Chrysler","Dodge","Dodge Brothers","Jeep","Kaiser Jeep","Ram","RAM","Rambler")
      if (any(Benzlist %in% Car)) {Benz[1] = 1;Benz[2] = length(which(Car %in% Benzlist))
      Benz[3:6] = as.numeric(summary(NStar[which(Car %in% Benzlist)]))[c(1,4,3,6)]}
    }
    if (Release >= "2008-04-03" & Release < "2008-08-07") # Land Rover & Jaguar to TATA 
    {
      TATAlist = c("Tata","Jaguar","Land-Rover")
      if (any(TATAlist %in% Car)) {TATA[1] = 1;TATA[2] = length(which(Car %in% TATAlist))
      TATA[3:6] = as.numeric(summary(NStar[which(Car %in% TATAlist)]))[c(1,4,3,6)]}
      
      FORDlist = c("Ford","Ford Shelby","Fordson","Lincoln")
      if (any(FORDlist %in% Car))  {FORD[1] = 1;FORD[2] = length(which(Car %in% FORDlist))
      FORD[3:6] = as.numeric(summary(NStar[which(Car %in% FORDlist)]))[c(1,4,3,6)]}
      
      GMlist = c("Pontiac","Cadillac","Cadillac-Gage","Chevrolet","Buick","McLaughlin-Buick","GM","GMC","WhiteGMC","GMDD","GM Allison")
      if (any(GMlist %in% Car)) {GM[1] = 1;GM[2] = length(which(Car %in% GMlist))
      GM[3:6] = as.numeric(summary(NStar[which(Car %in% GMlist)]))[c(1,4,3,6)]}
      
      HONDAlist = c("Honda","Acura")
      if (any(HONDAlist %in% Car)) {HONDA[1] = 1;HONDA[2] = length(which(Car %in% HONDAlist))
      HONDA[3:6] = as.numeric(summary(NStar[which(Car %in% HONDAlist)]))[c(1,4,3,6)]}
      
      Nissanlist = c("Nissan","Nissan Diesel","Infiniti")
      if (any(Nissanlist %in% Car)) {Nissan[1] = 1;Nissan[2] = length(which(Car %in% Nissanlist))
      Nissan[3:6] = as.numeric(summary(NStar[which(Car %in% Nissanlist)]))[c(1,4,3,6)]}
      
      Toyotalist = c("Toyota","Lexus")
      if (any(Toyotalist %in% Car)) {Toyota[1] = 1;Toyota[2] = length(which(Car %in% Toyotalist))
      Toyota[3:6] = as.numeric(summary(NStar[which(Car %in% Toyotalist)]))[c(1,4,3,6)]}
      
      # Fiatlist = c("Fiat","Fiatagri")
      # if (any(Fiatlist %in% Car)) {Fiat[1] = 1;Fiat[2] = length(which(Car %in% Fiatlist))
      # Fiat[3:6] = as.numeric(summary(NStar[which(Car %in% Fiatlist)]))[c(1,4,3,6)]}
      
      Benzlist = c("Daimler","Daimler-Benz","Mercedes-Benz","Chrysler","Dodge","Dodge Brothers","Jeep","Kaiser Jeep","Ram","RAM","Rambler")
      if (any(Benzlist %in% Car)) {Benz[1] = 1;Benz[2] = length(which(Car %in% Benzlist))
      Benz[3:6] = as.numeric(summary(NStar[which(Car %in% Benzlist)]))[c(1,4,3,6)]}
    }
    if (Release >= "2008-08-07" & Release < "2010-07-06") # No Nissan 
    {
      TATAlist = c("Tata","Jaguar","Land-Rover")
      if (any(TATAlist %in% Car)) {TATA[1] = 1;TATA[2] = length(which(Car %in% TATAlist))
      TATA[3:6] = as.numeric(summary(NStar[which(Car %in% TATAlist)]))[c(1,4,3,6)]}
      
      FORDlist = c("Ford","Ford Shelby","Fordson","Lincoln")
      if (any(FORDlist %in% Car))  {FORD[1] = 1;FORD[2] = length(which(Car %in% FORDlist))
      FORD[3:6] = as.numeric(summary(NStar[which(Car %in% FORDlist)]))[c(1,4,3,6)]}
      
      GMlist = c("Pontiac","Cadillac","Cadillac-Gage","Chevrolet","Buick","McLaughlin-Buick","GM","GMC","WhiteGMC","GMDD","GM Allison")
      if (any(GMlist %in% Car)) {GM[1] = 1;GM[2] = length(which(Car %in% GMlist))
      GM[3:6] = as.numeric(summary(NStar[which(Car %in% GMlist)]))[c(1,4,3,6)]}
      
      HONDAlist = c("Honda","Acura")
      if (any(HONDAlist %in% Car)) {HONDA[1] = 1;HONDA[2] = length(which(Car %in% HONDAlist))
      HONDA[3:6] = as.numeric(summary(NStar[which(Car %in% HONDAlist)]))[c(1,4,3,6)]}
      
      # Nissanlist = c("Nissan","Nissan Diesel","Infiniti")
      # if (any(Nissanlist %in% Car)) {Nissan[1] = 1;Nissan[2] = length(which(Car %in% Nissanlist))
      # Nissan[3:6] = as.numeric(summary(NStar[which(Car %in% Nissanlist)]))[c(1,4,3,6)]}
      
      Toyotalist = c("Toyota","Lexus")
      if (any(Toyotalist %in% Car)) {Toyota[1] = 1;Toyota[2] = length(which(Car %in% Toyotalist))
      Toyota[3:6] = as.numeric(summary(NStar[which(Car %in% Toyotalist)]))[c(1,4,3,6)]}
      
      # Fiatlist = c("Fiat","Fiatagri")
      # if (any(Fiatlist %in% Car)) {Fiat[1] = 1;Fiat[2] = length(which(Car %in% Fiatlist))
      # Fiat[3:6] = as.numeric(summary(NStar[which(Car %in% Fiatlist)]))[c(1,4,3,6)]}
      
      Benzlist = c("Daimler","Daimler-Benz","Mercedes-Benz","Chrysler","Dodge","Dodge Brothers","Jeep","Kaiser Jeep","Ram","RAM","Rambler")
      if (any(Benzlist %in% Car)) {Benz[1] = 1;Benz[2] = length(which(Car %in% Benzlist))
      Benz[3:6] = as.numeric(summary(NStar[which(Car %in% Benzlist)]))[c(1,4,3,6)]}
    }
    if (Release >= "2010-07-06" & Release < "2014-10-14") # No Benz
    {
      TATAlist = c("Tata","Jaguar","Land-Rover")
      if (any(TATAlist %in% Car)) {TATA[1] = 1;TATA[2] = length(which(Car %in% TATAlist))
      TATA[3:6] = as.numeric(summary(NStar[which(Car %in% TATAlist)]))[c(1,4,3,6)]}
      
      FORDlist = c("Ford","Ford Shelby","Fordson","Lincoln")
      if (any(FORDlist %in% Car))  {FORD[1] = 1;FORD[2] = length(which(Car %in% FORDlist))
      FORD[3:6] = as.numeric(summary(NStar[which(Car %in% FORDlist)]))[c(1,4,3,6)]}
      
      GMlist = c("Pontiac","Cadillac","Cadillac-Gage","Chevrolet","Buick","McLaughlin-Buick","GM","GMC","WhiteGMC","GMDD","GM Allison")
      if (any(GMlist %in% Car)) {GM[1] = 1;GM[2] = length(which(Car %in% GMlist))
      GM[3:6] = as.numeric(summary(NStar[which(Car %in% GMlist)]))[c(1,4,3,6)]}
      
      HONDAlist = c("Honda","Acura")
      if (any(HONDAlist %in% Car)) {HONDA[1] = 1;HONDA[2] = length(which(Car %in% HONDAlist))
      HONDA[3:6] = as.numeric(summary(NStar[which(Car %in% HONDAlist)]))[c(1,4,3,6)]}
      
      # Nissanlist = c("Nissan","Nissan Diesel","Infiniti")
      # if (any(Nissanlist %in% Car)) {Nissan[1] = 1;Nissan[2] = length(which(Car %in% Nissanlist))
      # Nissan[3:6] = as.numeric(summary(NStar[which(Car %in% Nissanlist)]))[c(1,4,3,6)]}
      
      Toyotalist = c("Toyota","Lexus")
      if (any(Toyotalist %in% Car)) {Toyota[1] = 1;Toyota[2] = length(which(Car %in% Toyotalist))
      Toyota[3:6] = as.numeric(summary(NStar[which(Car %in% Toyotalist)]))[c(1,4,3,6)]}
      
      # Fiatlist = c("Fiat","Fiatagri")
      # if (any(Fiatlist %in% Car)) {Fiat[1] = 1;Fiat[2] = length(which(Car %in% Fiatlist))
      # Fiat[3:6] = as.numeric(summary(NStar[which(Car %in% Fiatlist)]))[c(1,4,3,6)]}
      
      # Benzlist = c("Daimler","Daimler-Benz","Mercedes-Benz","Chrysler","Dodge","Dodge Brothers","Jeep","Kaiser Jeep","Ram","RAM","Rambler")
      # if (any(Benzlist %in% Car)) {Benz[1] = 1;Benz[2] = length(which(Car %in% Benzlist))
      # Benz[3:6] = as.numeric(summary(NStar[which(Car %in% Benzlist)]))[c(1,4,3,6)]}
    }
    if (Release >= "2014-10-14") # Fiat and Chrysler Back both belong to Fiat for FCA
    {
      TATAlist = c("Tata","Jaguar","Land-Rover")
      if (any(TATAlist %in% Car)) {TATA[1] = 1;TATA[2] = length(which(Car %in% TATAlist))
      TATA[3:6] = as.numeric(summary(NStar[which(Car %in% TATAlist)]))[c(1,4,3,6)]}
      
      FORDlist = c("Ford","Ford Shelby","Fordson","Lincoln")
      if (any(FORDlist %in% Car))  {FORD[1] = 1;FORD[2] = length(which(Car %in% FORDlist))
      FORD[3:6] = as.numeric(summary(NStar[which(Car %in% FORDlist)]))[c(1,4,3,6)]}
      
      GMlist = c("Pontiac","Cadillac","Cadillac-Gage","Chevrolet","Buick","McLaughlin-Buick","GM","GMC","WhiteGMC","GMDD","GM Allison")
      if (any(GMlist %in% Car)) {GM[1] = 1;GM[2] = length(which(Car %in% GMlist))
      GM[3:6] = as.numeric(summary(NStar[which(Car %in% GMlist)]))[c(1,4,3,6)]}
      
      HONDAlist = c("Honda","Acura")
      if (any(HONDAlist %in% Car)) {HONDA[1] = 1;HONDA[2] = length(which(Car %in% HONDAlist))
      HONDA[3:6] = as.numeric(summary(NStar[which(Car %in% HONDAlist)]))[c(1,4,3,6)]}
      
      # Nissanlist = c("Nissan","Nissan Diesel","Infiniti")
      # if (any(Nissanlist %in% Car)) {Nissan[1] = 1;Nissan[2] = length(which(Car %in% Nissanlist))
      # Nissan[3:6] = as.numeric(summary(NStar[which(Car %in% Nissanlist)]))[c(1,4,3,6)]}
      
      Toyotalist = c("Toyota","Lexus")
      if (any(Toyotalist %in% Car)) {Toyota[1] = 1;Toyota[2] = length(which(Car %in% Toyotalist))
      Toyota[3:6] = as.numeric(summary(NStar[which(Car %in% Toyotalist)]))[c(1,4,3,6)]}
      
      Fiatlist = c("Fiat","Fiatagri","Chrysler","Dodge","Dodge Brothers","Jeep","Kaiser Jeep","Ram","RAM","Rambler")
      if (any(Fiatlist %in% Car)) {Fiat[1] = 1;Fiat[2] = length(which(Car %in% Fiatlist))
      Fiat[3:6] = as.numeric(summary(NStar[which(Car %in% Fiatlist)]))[c(1,4,3,6)]}
      
      # Benzlist = c("Daimler","Daimler-Benz","Mercedes-Benz","Chrysler","Dodge","Dodge Brothers","Jeep","Kaiser Jeep","Ram","RAM","Rambler")
      # if (any(Benzlist %in% Car)) {Benz[1] = 1;Benz[2] = length(which(Car %in% Benzlist))
      # Benz[3:6] = as.numeric(summary(NStar[which(Car %in% Benzlist)]))[c(1,4,3,6)]}
    }
    return(c(Importdata[1,c(1:3,6:ncol(Importdata))],round(c(FORD,GM,HONDA,Nissan,Toyota,Fiat,TATA,Benz),4)))
  }
  
  {
    # 1. US 1st
    
    Vehicles1 = Vehicles[which(!is.na(Vehicles[,14])),]
    DummyData1 = as.data.frame(matrix(NA,length(unique(Vehicles1$MovieID)),52))
    Makes = c("Ford","GM","Honda","Nissan","Toyota","Fiat","TATA","Daimler-Benz")
    Varnames = list()
    for (i in 1:length(Makes))
    {
      Varnames[[i]] = paste0(Makes[i],c("_Dummy","_Count","_Min","_Mean","_Median","_Max"))
    }
    colnames(DummyData1) = c("MovieID","Movie","Year","US Earliest Release Date",as.character(unlist(Varnames)))
    
    for (i in 1:nrow(DummyData1))
    {
      DummyData1[i,] = matchcar(Vehicles1[which(Vehicles1$MovieID == unique(Vehicles1$MovieID)[i]),c(1:3,5,10,14)])
    }
    
    DummyData1 = merge(DummyData1[,c(1,4:ncol(DummyData1))],df_IMDB[,c(1:3)],by = "MovieID",all.y = T)
    DummyData1 = DummyData1[order(as.numeric(DummyData1$MovieID)),c(1,51,52,2:50)]
    DummyData1$MovieID = seq(1,nrow(DummyData1))
  }
  
  {
    # 1.1 US General
    
    Vehicles1_1 = Vehicles[which(!is.na(Vehicles[,15])),]
    DummyData1_1 = as.data.frame(matrix(NA,length(unique(Vehicles1_1$MovieID)),52))
    Makes = c("Ford","GM","Honda","Nissan","Toyota","Fiat","TATA","Daimler-Benz")
    Varnames = list()
    for (i in 1:length(Makes))
    {
      Varnames[[i]] = paste0(Makes[i],c("_Dummy","_Count","_Min","_Mean","_Median","_Max"))
    }
    colnames(DummyData1_1) = c("MovieID","Movie","Year","US General Date",as.character(unlist(Varnames)))
    
    for (i in 1:nrow(DummyData1_1))
    {
      DummyData1_1[i,] = matchcar(Vehicles1_1[which(Vehicles1_1$MovieID == unique(Vehicles1_1$MovieID)[i]),c(1:3,5,10,15)])
    }
    
    DummyData1_1 = merge(DummyData1_1[,c(1,4:ncol(DummyData1_1))],df_IMDB[,c(1:3)],by = "MovieID",all.y = T)
    DummyData1_1 = DummyData1_1[order(as.numeric(DummyData1_1$MovieID)),c(1,51,52,2:50)]
    DummyData1_1$MovieID = seq(1,nrow(DummyData1_1))
  }
  
  {
    # 1.2 US Festival
    
    Vehicles1_2 = Vehicles[which(!is.na(Vehicles[,16])),]
    DummyData1_2 = as.data.frame(matrix(NA,length(unique(Vehicles1_2$MovieID)),52))
    Makes = c("Ford","GM","Honda","Nissan","Toyota","Fiat","TATA","Daimler-Benz")
    Varnames = list()
    for (i in 1:length(Makes))
    {
      Varnames[[i]] = paste0(Makes[i],c("_Dummy","_Count","_Min","_Mean","_Median","_Max"))
    }
    colnames(DummyData1_2) = c("MovieID","Movie","Year","US Festival Date",as.character(unlist(Varnames)))
    
    for (i in 1:nrow(DummyData1_2))
    {
      DummyData1_2[i,] = matchcar(Vehicles1_2[which(Vehicles1_2$MovieID == unique(Vehicles1_2$MovieID)[i]),c(1:3,5,10,16)])
    }
    
    DummyData1_2 = merge(DummyData1_2[,c(1,4:ncol(DummyData1_2))],df_IMDB[,c(1:3)],by = "MovieID",all.y = T)
    DummyData1_2 = DummyData1_2[order(as.numeric(DummyData1_2$MovieID)),c(1,51,52,2:50)]
    DummyData1_2$MovieID = seq(1,nrow(DummyData1_2))
  }  
  
  {
    # 1.3 US Premiere
    
    Vehicles1_3 = Vehicles[which(!is.na(Vehicles[,18])),]
    DummyData1_3 = as.data.frame(matrix(NA,length(unique(Vehicles1_3$MovieID)),52))
    Makes = c("Ford","GM","Honda","Nissan","Toyota","Fiat","TATA","Daimler-Benz")
    Varnames = list()
    for (i in 1:length(Makes))
    {
      Varnames[[i]] = paste0(Makes[i],c("_Dummy","_Count","_Min","_Mean","_Median","_Max"))
    }
    colnames(DummyData1_3) = c("MovieID","Movie","Year","US Premiere Date",as.character(unlist(Varnames)))
    
    for (i in 1:nrow(DummyData1_3))
    {
      DummyData1_3[i,] = matchcar(Vehicles1_3[which(Vehicles1_3$MovieID == unique(Vehicles1_3$MovieID)[i]),c(1:3,5,10,18)])
    }
    
    DummyData1_3 = merge(DummyData1_3[,c(1,4:ncol(DummyData1_3))],df_IMDB[,c(1:3)],by = "MovieID",all.y = T)
    DummyData1_3 = DummyData1_3[order(as.numeric(DummyData1_3$MovieID)),c(1,51,52,2:50)]
    DummyData1_3$MovieID = seq(1,nrow(DummyData1_3))
  }  
  
  {
    # 1.4 US Limited
    
    Vehicles1_4 = Vehicles[which(!is.na(Vehicles[,20])),]
    DummyData1_4 = as.data.frame(matrix(NA,length(unique(Vehicles1_4$MovieID)),52))
    Makes = c("Ford","GM","Honda","Nissan","Toyota","Fiat","TATA","Daimler-Benz")
    Varnames = list()
    for (i in 1:length(Makes))
    {
      Varnames[[i]] = paste0(Makes[i],c("_Dummy","_Count","_Min","_Mean","_Median","_Max"))
    }
    colnames(DummyData1_4) = c("MovieID","Movie","Year","US Limited Date",as.character(unlist(Varnames)))
    
    for (i in 1:nrow(DummyData1_4))
    {
      DummyData1_4[i,] = matchcar(Vehicles1_4[which(Vehicles1_4$MovieID == unique(Vehicles1_4$MovieID)[i]),c(1:3,5,10,20)])
    }
    
    DummyData1_4 = merge(DummyData1_4[,c(1,4:ncol(DummyData1_4))],df_IMDB[,c(1:3)],by = "MovieID",all.y = T)
    DummyData1_4 = DummyData1_4[order(as.numeric(DummyData1_4$MovieID)),c(1,51,52,2:50)]
    DummyData1_4$MovieID = seq(1,nrow(DummyData1_4))
  }  
  
  {
    # 1.5 US Internet
    
    Vehicles1_5 = Vehicles[which(!is.na(Vehicles[,21])),]
    DummyData1_5 = as.data.frame(matrix(NA,length(unique(Vehicles1_5$MovieID)),52))
    Makes = c("Ford","GM","Honda","Nissan","Toyota","Fiat","TATA","Daimler-Benz")
    Varnames = list()
    for (i in 1:length(Makes))
    {
      Varnames[[i]] = paste0(Makes[i],c("_Dummy","_Count","_Min","_Mean","_Median","_Max"))
    }
    colnames(DummyData1_5) = c("MovieID","Movie","Year","US Internet Date",as.character(unlist(Varnames)))
    
    for (i in 1:nrow(DummyData1_5))
    {
      DummyData1_5[i,] = matchcar(Vehicles1_5[which(Vehicles1_5$MovieID == unique(Vehicles1_5$MovieID)[i]),c(1:3,5,10,21)])
    }
    
    DummyData1_5 = merge(DummyData1_5[,c(1,4:ncol(DummyData1_5))],df_IMDB[,c(1:3)],by = "MovieID",all.y = T)
    DummyData1_5 = DummyData1_5[order(as.numeric(DummyData1_5$MovieID)),c(1,51,52,2:50)]
    DummyData1_5$MovieID = seq(1,nrow(DummyData1_5))
  }  
  
  {  
    # 2. Global 1st
    
    Vehicles2 = Vehicles[which(!is.na(Vehicles[,24])),]
    DummyData2 = as.data.frame(matrix(NA,length(unique(Vehicles2$MovieID)),53))
    colnames(DummyData2) = c("MovieID","Movie","Year","Global Earliest Release Country","Global Earliest Release Date",
                             as.character(unlist(Varnames)))
    
    for (i in 1:nrow(DummyData2))
    {
      DummyData2[i,] = matchcar(Vehicles2[which(Vehicles2$MovieID == unique(Vehicles2$MovieID)[i]),c(1:3,5,10,23,24)])
    }
    DummyData2 = DummyData2[order(as.numeric(DummyData2$MovieID)),]
    {  
      # if (C == "Hong Kong" || c == "Macao") TZ = "Asia/Hong_Kong"
      

      # Clist = c("France","Russia","US","Australia","UK","Canada","Denmark","New Zealand","Brazil","Mexico","Chile","Indonesia",
      #           "Kiribati","Democratic Republic of the Congo","Ecuador","Federated States of Micronesia","Kazakhstan",
      #           "Kingdom of the Netherlands","Mongolia","Papua New Guinea","Portugal","South Africa","Spain","Ukraine")
      # Countries = unique(c(Vehicles[,c(23)],Vehicles[,25],Vehicles[,27],Vehicles[,30],Vehicles[,33],Vehicles[,35])
      #                    [which(!is.na(c(Vehicles[,c(23)],Vehicles[,25],Vehicles[,27],Vehicles[,30],Vehicles[,33],
      #                                    Vehicles[,35])))])
      # Clist[which(Clist %in% Countries)]
      # [1] "France"       "Russia"       "Australia"    "UK"           "Canada"       "Denmark"      "New Zealand"  "Brazil"
      # [9] "Mexico"       "Chile"        "Indonesia"    "Kazakhstan"   "Portugal"     "South Africa" "Spain"        "Ukraine"
      
      # Function to convert timezone (Match the capital time)
    }
    DummyData2[,5] = paste0(DummyData2[,5]," 12:00:01")
    
    CountrytoTZ = function(ImportData)
    {
      {C = ImportData[1]
      if (C == "Norway") TZ = "Europe/Oslo"
      if (C == "Lebanon") TZ = "Asia/Beirut"
      if (C == "Poland") TZ = "Europe/Warsaw"
      if (C == "Latvia") TZ = "Europe/Riga"
      if (C == "Fiji") TZ = "Pacific/Fiji"
      if (C == "Georgia") TZ = "Asia/Tbilisi"
      if (C == "Bosnia and Herzegovina") TZ = "Europe/Sarajevo"
      if (C == "Aruba") TZ = "America/Aruba"
      if (C == "Turkey") TZ = "Turkey"
      if (C == "Bermuda") TZ = "Atlantic/Bermuda"
      if (C == "Dominican Republic") TZ = "America/Santo_Domingo"
      if (C == "Bahamas") TZ = "America/Nassau"
      if (C == "Estonia") TZ = "Europe/Tallinn"
      if (C == "Peru") TZ = "America/Lima"
      if (C == "Iraq") TZ = "Asia/Baghdad"
      if (C == "Portugal") TZ = "Europe/Lisbon"
      if (C == "South Africa") TZ = "Africa/Johannesburg"
      if (C == "USA") TZ = "America/New_York"
      if (C == "UK") TZ = "Europe/London"
      if (C == "Serbia and Montenegro") TZ = "Europe/Belgrade"
      if (C == "Netherlands Antilles") TZ = "America/Puerto_Rico"
      if (C == "Australia") TZ = "Australia/Canberra"
      if (C == "France") TZ = "Europe/Paris"
      if (C == "Russia") TZ = "Europe/Moscow"
      if (C == "Canada") TZ = "Canada/Eastern"
      if (C == "Denmark") TZ = "Europe/Copenhagen"
      if (C == "New Zealand") TZ = "Pacific/Auckland"
      if (C == "Brazil") TZ = "Brazil/East"
      if (C == "Mexico" || C == "El Salvador") TZ = "America/Mexico_City"
      if (C == "Chile") TZ = "Chile/Continental"
      if (C == "Indonesia") TZ = "Asia/Jakarta"
      if (C == "Kazakhstan") TZ = "Asia/Almaty"
      if (C == "Spain") TZ = "Europe/Madrid"
      if (C == "Ukraine") TZ = "Europe/Kiev"
      if (C == "Argentina") TZ = "America/Argentina/Buenos_Aires"
      if (C == "Puerto Rico") TZ = "America/Puerto_Rico"
      if (C == "Netherlands") TZ = "Europe/Amsterdam"
      if (C == "Israel") TZ = "Israel"
      if (C == "United Arab Emirates") TZ = "Asia/Dubai"
      if (C == "Japan") TZ = "Japan"
      if (C == "Kuwait" || C == "Qatar") TZ = "Asia/Kuwait"
      if (C == "Jamaica") TZ = "America/Jamaica"
      if (C == "Switzerland") TZ = "Europe/Zurich"
      if (C == "Germany") TZ = "Europe/Berlin"
      if (C == "Hong Kong" || C == "Macao") TZ = "Asia/Hong_Kong"
      if (C == "Italy") TZ = "Europe/Rome"
      if (C == "Romania") TZ = "Europe/Bucharest"
      if (C == "Thailand") TZ = "Asia/Bangkok"
      if (C == "Singapore") TZ = "Asia/Singapore"
      if (C == "Philippines") TZ = "Asia/Manila"
      if (C == "Sweden") TZ = "Europe/Stockholm"
      if (C == "Trinidad and Tobago") TZ = "America/Port_of_Spain"
      if (C == "Greece") TZ = "Europe/Athens"
      if (C == "Taiwan") TZ = "Asia/Taipei"
      if (C == "Malaysia") TZ = "Asia/Kuala_Lumpur"
      if (C == "South Korea") TZ = "Asia/Seoul"
      if (C == "Uruguay") TZ = "America/Montevideo"
      if (C == "Belgium") TZ = "Europe/Brussels"
      if (C == "Czech Republic" || C == "Slovakia") TZ = "Europe/Prague"
      if (C == "Iceland") TZ = "Atlantic/Reykjavik"
      if (C == "Colombia") TZ = "America/Bogota"
      if (C == "Bolivia") TZ = "America/La_Paz"
      if (C == "Croatia") TZ = "Europe/Zagreb"
      if (C == "China") TZ = "Asia/Shanghai"
      if (C == "Egypt") TZ = "Africa/Cairo"
      if (C == "Hungary") TZ = "Europe/Budapest"
      if (C == "Bulgaria") TZ = "Europe/Sofia"
      if (C == "Austria") TZ = "Europe/Vienna"
      if (C == "Finland") TZ = "Europe/Helsinki"
      if (C == "Lithuania") TZ = "Europe/Vilnius"
      if (C == "Bahrain") TZ = "Asia/Bahrain"
      if (C == "Belarus") TZ = "Europe/Minsk"
      if (C == "Slovenia") TZ = "Europe/Ljubljana"
      if (C == "Armenia") TZ = "Asia/Yerevan"
      if (C == "Azerbaijan") TZ = "Asia/Baku"
      if (C == "India" || C == "Sri Lanka") TZ = "Asia/Kolkata"
      if (C == "Cambodia") TZ = "Asia/Phnom_Penh"
      if (C == "Albania") TZ = "Europe/Tirane"
      if (C == "Ireland") TZ = "Europe/Dublin"}
      # else TZ = TimeZonelist[which(TimeZonelist$Country.Name == as.character(C)),3][1]
      LocalTime = as.POSIXct(as.character(ImportData[2]), tz=TZ)
      as.POSIXct(as.character(ImportData[2]), tz="CST6CDT")
      NYTime = format(LocalTime,tz="America/New_York")
      return(NYTime)
    }
    
    DummyData2$'Equivalent NY Time' = NA
    for (i in 1:nrow(DummyData2))
    {
      DummyData2$'Equivalent NY Time'[i] = CountrytoTZ(DummyData2[i,c(4,5)])
    }
    DummyData2 = DummyData2[,c(1:5,ncol(DummyData2),6:(ncol(DummyData2)-1))]
    DummyData2 = merge(DummyData2[,c(1,4:ncol(DummyData2))],df_IMDB[,c(1:3)],by = "MovieID",all.y = T)
    DummyData2 = DummyData2[order(as.numeric(DummyData2$MovieID)),c(1,53,54,2:52)]
    DummyData2$MovieID = seq(1,nrow(DummyData2))
  }
  
  {  
    # 2.1 Global General
    
    Vehicles2_1 = Vehicles[which(!is.na(Vehicles[,26])),]
    DummyData2_1 = as.data.frame(matrix(NA,length(unique(Vehicles2_1$MovieID)),53))
    colnames(DummyData2_1) = c("MovieID","Movie","Year","Global General Country","Global General Date",
                               as.character(unlist(Varnames)))
    
    for (i in 1:nrow(DummyData2_1))
    {
      DummyData2_1[i,] = matchcar(Vehicles2_1[which(Vehicles2_1$MovieID == unique(Vehicles2_1$MovieID)[i]),c(1:3,5,10,25,26)])
    }
    DummyData2_1 = DummyData2_1[order(as.numeric(DummyData2_1$MovieID)),]
    DummyData2_1[,5] = paste0(DummyData2_1[,5]," 12:00:01")
    
    DummyData2_1$'Equivalent NY Time' = NA
    for (i in 1:nrow(DummyData2_1))
    {
      DummyData2_1$'Equivalent NY Time'[i] = tryCatch(CountrytoTZ(DummyData2_1[i,c(4,5)]),error = function(e) NA)
    }
    DummyData2_1 = DummyData2_1[,c(1:5,ncol(DummyData2_1),6:(ncol(DummyData2_1)-1))]
    DummyData2_1 = merge(DummyData2_1[,c(1,4:ncol(DummyData2_1))],df_IMDB[,c(1:3)],by = "MovieID",all.y = T)
    DummyData2_1 = DummyData2_1[order(as.numeric(DummyData2_1$MovieID)),c(1,53,54,2:52)]
    DummyData2_1$MovieID = seq(1,nrow(DummyData2_1))
  }
  
  {  
    # 2.2 Global Festival
    
    Vehicles2_2 = Vehicles[which(!is.na(Vehicles[,29])),]
    DummyData2_2 = as.data.frame(matrix(NA,length(unique(Vehicles2_2$MovieID)),53))
    colnames(DummyData2_2) = c("MovieID","Movie","Year","Global Festival Country","Global Festival Date",
                               as.character(unlist(Varnames)))
    
    for (i in 1:nrow(DummyData2_2))
    {
      DummyData2_2[i,] = matchcar(Vehicles2_2[which(Vehicles2_2$MovieID == unique(Vehicles2_2$MovieID)[i]),c(1:3,5,10,27,29)])
    }
    DummyData2_2 = DummyData2_2[order(as.numeric(DummyData2_2$MovieID)),]
    DummyData2_2[,5] = paste0(DummyData2_2[,5]," 12:00:01")
    
    DummyData2_2$'Equivalent NY Time' = NA
    for (i in 1:nrow(DummyData2_2))
    {
      DummyData2_2$'Equivalent NY Time'[i] = tryCatch(CountrytoTZ(DummyData2_2[i,c(4,5)]),error = function(e) NA)
    }
    DummyData2_2 = DummyData2_2[,c(1:5,ncol(DummyData2_2),6:(ncol(DummyData2_2)-1))]
    DummyData2_2 = merge(DummyData2_2[,c(1,4:ncol(DummyData2_2))],df_IMDB[,c(1:3)],by = "MovieID",all.y = T)
    DummyData2_2 = DummyData2_2[order(as.numeric(DummyData2_2$MovieID)),c(1,53,54,2:52)]
    DummyData2_2$MovieID = seq(1,nrow(DummyData2_2))
  }  
  
  {  
    # 2.3 Global Premiere
    
    Vehicles2_3 = Vehicles[which(!is.na(Vehicles[,32])),]
    DummyData2_3 = as.data.frame(matrix(NA,length(unique(Vehicles2_3$MovieID)),53))
    colnames(DummyData2_3) = c("MovieID","Movie","Year","Global Premieree Country","Global Premiere Date",
                               as.character(unlist(Varnames)))
    
    for (i in 1:nrow(DummyData2_3))
    {
      DummyData2_3[i,] = matchcar(Vehicles2_3[which(Vehicles2_3$MovieID == unique(Vehicles2_3$MovieID)[i]),c(1:3,5,10,30,32)])
    }
    DummyData2_3 = DummyData2_3[order(as.numeric(DummyData2_3$MovieID)),]
    DummyData2_3[,5] = paste0(DummyData2_3[,5]," 12:00:01")
    
    DummyData2_3$'Equivalent NY Time' = NA
    for (i in 1:nrow(DummyData2_3))
    {
      DummyData2_3$'Equivalent NY Time'[i] = tryCatch(CountrytoTZ(DummyData2_3[i,c(4,5)]),error = function(e) NA)
    }
    DummyData2_3 = DummyData2_3[,c(1:5,ncol(DummyData2_3),6:(ncol(DummyData2_3)-1))]
    DummyData2_3 = merge(DummyData2_3[,c(1,4:ncol(DummyData2_3))],df_IMDB[,c(1:3)],by = "MovieID",all.y = T)
    DummyData2_3 = DummyData2_3[order(as.numeric(DummyData2_3$MovieID)),c(1,53,54,2:52)]
    DummyData2_3$MovieID = seq(1,nrow(DummyData2_3))
  }   
  
  {  
    # 2.4 Global Limited
    
    Vehicles2_4 = Vehicles[which(!is.na(Vehicles[,34])),]
    DummyData2_4 = as.data.frame(matrix(NA,length(unique(Vehicles2_4$MovieID)),53))
    colnames(DummyData2_4) = c("MovieID","Movie","Year","Global Limited Country","Global Limited Date",
                               as.character(unlist(Varnames)))
    
    for (i in 1:nrow(DummyData2_4))
    {
      DummyData2_4[i,] = matchcar(Vehicles2_4[which(Vehicles2_4$MovieID == unique(Vehicles2_4$MovieID)[i]),c(1:3,5,10,33,34)])
    }
    DummyData2_4 = DummyData2_4[order(as.numeric(DummyData2_4$MovieID)),]
    DummyData2_4[,5] = paste0(DummyData2_4[,5]," 12:00:01")
    
    DummyData2_4$'Equivalent NY Time' = NA
    for (i in 1:nrow(DummyData2_4))
    {
      DummyData2_4$'Equivalent NY Time'[i] = tryCatch(CountrytoTZ(DummyData2_4[i,c(4,5)]),error = function(e) NA)
    }
    DummyData2_4 = DummyData2_4[,c(1:5,ncol(DummyData2_4),6:(ncol(DummyData2_4)-1))]
    DummyData2_4 = merge(DummyData2_4[,c(1,4:ncol(DummyData2_4))],df_IMDB[,c(1:3)],by = "MovieID",all.y = T)
    DummyData2_4 = DummyData2_4[order(as.numeric(DummyData2_4$MovieID)),c(1,53,54,2:52)]
    DummyData2_4$MovieID = seq(1,nrow(DummyData2_4))
  }   
  
  {  
    # 2.5 Global Internet
    
    Vehicles2_5 = Vehicles[which(!is.na(Vehicles[,36])),]
    DummyData2_5 = as.data.frame(matrix(NA,length(unique(Vehicles2_5$MovieID)),53))
    colnames(DummyData2_5) = c("MovieID","Movie","Year","Global Internet Country","Global Internet Date",
                               as.character(unlist(Varnames)))
    
    for (i in 1:nrow(DummyData2_5))
    {
      DummyData2_5[i,] = matchcar(Vehicles2_5[which(Vehicles2_5$MovieID == unique(Vehicles2_5$MovieID)[i]),c(1:3,5,10,35,36)])
    }
    DummyData2_5 = DummyData2_5[order(as.numeric(DummyData2_5$MovieID)),]
    DummyData2_5[,5] = paste0(DummyData2_5[,5]," 12:00:01")
    
    DummyData2_5$'Equivalent NY Time' = NA
    for (i in 1:nrow(DummyData2_5))
    {
      DummyData2_5$'Equivalent NY Time'[i] = tryCatch(CountrytoTZ(DummyData2_5[i,c(4,5)]),error = function(e) NA)
    }
    DummyData2_5 = DummyData2_5[,c(1:5,ncol(DummyData2_5),6:(ncol(DummyData2_5)-1))]
    DummyData2_5 = merge(DummyData2_5[,c(1,4:ncol(DummyData2_5))],df_IMDB[,c(1:3)],by = "MovieID",all.y = T)
    DummyData2_5 = DummyData2_5[order(as.numeric(DummyData2_5$MovieID)),c(1,53,54,2:52)]
    DummyData2_5$MovieID = seq(1,nrow(DummyData2_5))
  }     
  
}


# 3. Boxoffice Data

{
  df_IMDB$MovieID = as.numeric(df_IMDB$MovieID)
  BoxofficeData = read.xlsx("Boxoffice All Movies (90-18).xlsx",1)
  
  # Remove dup movies
  
  BoxofficeData = BoxofficeData[which(BoxofficeData$MovieID %in% ValidMovieID),]
  
  Merged2 = merge(BoxofficeData[,c(1,5:ncol(BoxofficeData))], df_IMDB, by = "MovieID", all = T)
  names(Merged2) = gsub("\\.", " ", names(Merged2))
  Merged2 = Merged2[order(as.numeric(Merged2$MovieID)),]
}


# 4. Blu-ray Data

{
  Blueray = read.xlsx("Blueray All Movies (90-18).xlsx",1)
  
  # Remove dup movies
  
  Blueray = Blueray[which(Blueray$MovieID %in% ValidMovieID),]
  
  
  # all.equal(Blueray$MovieID, Merged2$MovieID)
  # Merged3 = merge(Blueray[,c(1,4:7)], Merged2[unique(Merged2$MovieID),], by = "MovieID",all = T)
  Merged3 = cbind(Blueray[,c(1,4:7)], Merged2[,c(2:ncol(Merged2))])
  names(Merged3) = gsub("\\.", " ", names(Merged3))
}


# 5. 3D Data

{
  # # merge w former data
  # FormerData = read.xlsx("Full Movie Data 90-04 (Modified 0303).xlsx", 2)
  # names(FormerData) = gsub("\\.", " ", names(FormerData))
  # for (i in 1:nrow(FormerData))
  # {
  #   if (Movie3D$Movie[i] %in% FormerData$MOVNAME) print(Movie3D$Movie[i])
  # }  
  # Newmerged4 = merge(FormerData[,-c(16:20)], Movie3D, by.x = "MOVNAME", by.y = "Movie", all.x = T)
  # Newmerged4 = Newmerged4[,c(2,1,3:15,49:54,16:48)]
  # Newmerged4 = Newmerged4[order(as.numeric(Newmerged4$MOVIEID)),]
  # write.xlsx(Newmerged4, "Movie with 3D Info.xlsx")
  
  Movie3D = read.xlsx("Cleaned 3D Movie List.xlsx",1)
  for (i in 1:nrow(Movie3D))
  {
    if (Movie3D$Movie[i] %in% df_IMDB$Movie | Movie3D$Movie[i] %in% df_IMDB$`Name on IMDB` | Movie3D$Movie[i] %in% df_IMDB$`Other Name`) print(Movie3D$Movie[i])
  }
  Movie3D$`3D.Release.Date` = substr(Movie3D$`3D.Release.Date`, 1,4)
  
  Merged4 = merge(Merged3, Movie3D, by.x = c("Movie","Year"), by.y = c("Movie","3D.Release.Date"), all.x = T)
}


# 6. Reform Data

{
  FullMovieData = Merged4[,c(3,1,2,16,15,23,8:14,4:7,56:60,17:22,24:50,53:55)]
  names(FullMovieData)[c(1:6,14,16,23:32,40:42)] = 
    c("MOVIEID","MOVNAME","YEAR","DUMSEQ","DUMMARVEL","MOVBUDG","BLURAYRLDATE","DVDRELDATE","CRITRAT",
      "CLASSIF","RUNTIME","MOVGENRES","DUMACTION","DUMCOM","OSCARAWRD","OSCARNOM","US Earliest Release Type",
      "US Earliest Release Date","Global Earliest Release Type","Global Earliest Release Country",
      "Global Earliest Release Date")
  FullMovieData = FullMovieData[order(as.numeric(FullMovieData$MOVIEID)),]
  FullMovieData = unique(FullMovieData);names(FullMovieData) = gsub("\\.", " ", names(FullMovieData))
  names(FullMovieData) = gsub("Festivel", "Festival", names(FullMovieData))
  names(FullMovieData)[c(26,27,29,31,32,37,40,43,45,47)+7] = paste0(names(FullMovieData)[c(26,27,29,31,32,37,40,43,45,47)+7]," Date")
  
  # Add VHS movie data (Renamed)
  
  VHS = read.xlsx("other media format_1st sample example.xlsx",1)
  VUS_list = unique(VHS[c("MOVNAME", "VHS.Release.Date")])
  DatawithVHS = merge(FullMovieData, VUS_list, by = "MOVNAME",all.x = T)
  names(DatawithVHS) = gsub("\\.", " ", names(DatawithVHS))
  DatawithVHS = DatawithVHS[,c(1:13,59,14:58)]
  DatawithVHS$`VHS Release Date` = as.Date(as.character(DatawithVHS$`VHS Release Date`), "%Y%m%d")
  DatawithVHS = DatawithVHS[order(as.numeric(DatawithVHS$MOVIEID)),]
  DatawithVHS = DatawithVHS[,c(2,1,3:ncol(DatawithVHS))]
}


# 7. Add seqIDs

{
  # Alldata = read.xlsx("Full Movie Data, All Movies (90-18).xlsx",2)
  # names(Alldata) = gsub("\\.", " ", names(Alldata))
  
  # seqmovies = read.xlsx("Movies with sequels (all).xlsx",1)
  # # DatawithVHS
  # seqmovies$Included = 0
  # seqmovies$Included[which(seqmovies$URL %in% DatawithVHS$`IMDB URL`)] = 1
  # write.xlsx(seqmovies,"Movies with sequels (all).xlsx")
  
  # Filtered movies
  Filtered = read.xlsx("Filted Movies with sequels (all).xlsx",1)
  
  Filtered$Sequence[1] = 1
  i = 2
  while (i <= nrow(Filtered))
  {
    if (Filtered$SeqName[i-1] == Filtered$SeqName[i]) Filtered$Sequence[i] = Filtered$Sequence[i-1] + 1
    else Filtered$Sequence[i] = 1
    i = i + 1
  }
  write.xlsx(Filtered,"Cleaned sequel movies (all dates available with other media types).xlsx")
  
  # Add the filtered sequel movies to the original data set
  
  # OriginalMovieData = read.xlsx("Full Movie Data, All Movies (90-18).xlsx",2)
  merged5 = merge(DatawithVHS,Filtered, by.x = "IMDB URL", by.y = "URL", all.x = T)
  MovieDataWSeq = merged5[,c(2:4,62,63,5:56,1,57:59)]
  # length(table(MovieDataWSeq$SeqName)) # 327 movie series
  # Then if there's only 1 movie in the series, remove it
  MovieDataWSeq$New_DUMSEQ = MovieDataWSeq$DUMSEQ
  MovieDataWSeq$New_Sequence = MovieDataWSeq$Sequence
  MovieDataWSeq$New_SeqName = MovieDataWSeq$SeqName
  
  for (i in 1:length(table(MovieDataWSeq$SeqName)))
  {
    if (table(MovieDataWSeq$SeqName)[i] == 1)
      {
        MovieDataWSeq$New_DUMSEQ[which(MovieDataWSeq$SeqName == names(table(MovieDataWSeq$SeqName)[i]))] = 0
        MovieDataWSeq$New_Sequence[which(MovieDataWSeq$SeqName == names(table(MovieDataWSeq$SeqName)[i]))] = NA
        MovieDataWSeq$New_SeqName[which(MovieDataWSeq$SeqName == names(table(MovieDataWSeq$SeqName)[i]))] = NA
      }
  }
  
  MovieDataWSeq = MovieDataWSeq[,c(1:3,62:64,7:61)]
  MovieDataWSeq$DUMMARVEL[which(MovieDataWSeq$DUMMARVEL != "1")] = "0";MovieDataWSeq$DUMMARVEL[which(is.na(MovieDataWSeq$DUMMARVEL))] = "0"
  names(MovieDataWSeq)[c(4,5,6)] = c("DUMSEQ","Sequence","SeqName")
  # write.xlsx(MovieDataWSeq[order(as.numeric(MovieDataWSeq$MOVIEID)),],"Movies with SEQ var added.xlsx")
  
  MovieDataWSeq = MovieDataWSeq[order(MovieDataWSeq$YEAR),]
  MovieDataWSeq = MovieDataWSeq[order(MovieDataWSeq$SeqName),]
  
  
  # MovieDataWSeq = read.xlsx("Movies with SEQ var added.xlsx")
  MovieDataWSeq$SequenceInData = NA
  MovieDataWSeq$SequenceInData[1] = 1
  i = 2
  while (i <= nrow(MovieDataWSeq))
  {
    if (MovieDataWSeq$SeqName[i-1] == MovieDataWSeq$SeqName[i]) MovieDataWSeq$SequenceInData[i] = MovieDataWSeq$SequenceInData[i-1] + 1
    else MovieDataWSeq$SequenceInData[i] = 1
    i = i + 1
  }
  MovieDataWSeq = MovieDataWSeq[,c(1:5,62,6:61)]
  write.xlsx(MovieDataWSeq[order(as.numeric(MovieDataWSeq$MOVIEID)),],"Movies with SEQ var to replace the movie info.xlsx")
}


# Write Data

{
  MovieDataWSeq = MovieDataWSeq[order(as.numeric(MovieDataWSeq$MOVIEID)),]
  MovieDataWSeq$MOVIEID = seq(1,nrow(MovieDataWSeq))
  wb <- createWorkbook()
  addWorksheet(wb, sheetName = "Data (90-18)")
  addWorksheet(wb, sheetName = "Auto Data Using US Earliest")
  addWorksheet(wb, sheetName = "Auto Data Using US General")
  addWorksheet(wb, sheetName = "Auto Data Using US Festival")
  addWorksheet(wb, sheetName = "Auto Data Using US Premiere")
  addWorksheet(wb, sheetName = "Auto Data Using US Limited")
  addWorksheet(wb, sheetName = "Auto Data Using US Internet")
  addWorksheet(wb, sheetName = "Auto Data Using Global Earliest")
  addWorksheet(wb, sheetName = "Auto Data Using Global General")
  addWorksheet(wb, sheetName = "Auto Data Using Global Festival")
  addWorksheet(wb, sheetName = "Auto Data Using Global Premiere")
  addWorksheet(wb, sheetName = "Auto Data Using Global Limited")
  addWorksheet(wb, sheetName = "Auto Data Using Global Internet")
  
  writeData(wb, "Data (90-18)", MovieDataWSeq)
  writeData(wb, "Auto Data Using US Earliest", DummyData1)
  writeData(wb, "Auto Data Using US General", DummyData1_1)
  writeData(wb, "Auto Data Using US Festival", DummyData1_2)
  writeData(wb, "Auto Data Using US Premiere", DummyData1_3)
  writeData(wb, "Auto Data Using US Limited", DummyData1_4)
  writeData(wb, "Auto Data Using US Internet", DummyData1_5)
  writeData(wb, "Auto Data Using Global Earliest", DummyData2)
  writeData(wb, "Auto Data Using Global General", DummyData2_1)
  writeData(wb, "Auto Data Using Global Festival", DummyData2_2)
  writeData(wb, "Auto Data Using Global Premiere", DummyData2_3)
  writeData(wb, "Auto Data Using Global Limited", DummyData2_4)
  writeData(wb, "Auto Data Using Global Internet", DummyData2_5)
  saveWorkbook(wb, "Full Movie Data, All Movies (90-18).xlsx", overwrite = T)
  
}







