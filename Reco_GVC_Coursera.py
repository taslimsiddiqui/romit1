# -*- coding: utf-8 -*-
"""
Created on Mon Jun 10 14:17:22 2024

@author: romitbenkar
"""

import os

# Specify the path to the new working directory
new_directory = "D:\Python_Folder"

# Change the current working directory
os.chdir(new_directory)

# Verify the change
print("Current Working Directory: ", os.getcwd())

"Recommendation"

import pandas as pd 
import numpy as np 
GVC= pd.read_excel("C:\\Users\\Lenovo\\Downloads\\Wizr dump.xlsx",sheet_name='HR') 
# GVC.info() 
# GVC.isnull().sum() 

  
GVC['Course Syllabus'] = GVC['Course Syllabus'].fillna(0) 
GVC['About Course'] = GVC['About Course'].fillna(0) 
GVC['New_Concepts'] = GVC['New_Concepts'].fillna(0) 
GVC['Course Name'] = GVC['Course Name'].astype(str) 
GVC['About Course'] = GVC['About Course'].astype(str) 
GVC['Course Syllabus'] = GVC['Course Syllabus'].astype(str) 
GVC['Tags']=GVC['Course Name']+GVC['About Course']+GVC['Course Syllabus'] 

# print(GVC['Tags']) 

new_df=GVC[['Course Name','New_Concepts','Tags',]] 

new_df.dropna(inplace=True) 

# new_df.isnull().sum() 

new_df['Tags']=new_df['Tags'].apply(lambda x:x.lower()) 

# print(new_df) 

  

import nltk 

from nltk.stem.porter import PorterStemmer 

ps = PorterStemmer() 

  

def stem(text): 

    y=[] 

    for i in text.split(): 

        y.append(ps.stem(i)) 

    return" ".join(y) 

  

  

new_df['Tags']=new_df['Tags'].apply(stem) 

# print(new_df['Tags']) 

# new_df.head() 

  

from sklearn.feature_extraction.text import CountVectorizer 

cv = CountVectorizer(max_features=15000,stop_words='english') 

  

vectors=cv.fit_transform(new_df['Tags']).toarray() 

vectors[0] 

feature_names=cv.get_feature_names_out() 

for feature in feature_names: 

    print(feature) 

  

from sklearn.metrics.pairwise import cosine_similarity 

similarity=cosine_similarity(vectors) 

similarity[2] 

  

# new_df.info() 

# new_df.head() 

#to remove unnecesary spaces. 

new_df['Course Name'] = new_df['Course Name'].str.strip() 

new_df[new_df['Course Name']=='MASTER OF BUSINESS ADMINISTRATION - ENGINEERING MANAGEMENT'].index[0] 

def recommend(SSC): 

    course_index = new_df[new_df['Course Name'] == SSC].index[0] 

    distances = similarity[course_index] 

    course_list = sorted(list(enumerate(distances)), reverse=True, key=lambda x: x[1])[1:6] 

  

    recommended_concepts = [] 

    for i in course_list: 

        recommended_concepts.append(new_df.iloc[i[0]].New_Concepts) 

     

    return recommended_concepts 

  

  

GVC['Recommendation']=GVC['Course Name'].apply(recommend) 

GVC['Recommendation'].head() 

  

file_path = "D:\\OneDrive_2024-06-01\\ABC Catalogue Sharing\\Course_Reco_Output.xlsx" 

GVC['Recommendation'].to_excel(file_path) 

print("Data has been written to the Excel file successfully.") 

# create new excel/ create new column if excel is already exist 

#already create a coumn header name 

#open excel sheet to write the data 

# append the data which coming fromm line no 72 

#%% 

import pandas as pd 

import numpy as np 

GVC= pd.read_excel("D://Python_Folder//Concept_reco.xlsx",sheet_name='Sheet1') 

GVC.info() 

GVC.isnull().sum() 

GVC['Concept_Description'] = GVC['Concept_Description'].fillna(0) 

GVC['Concept_name'] = GVC['Concept_name'].fillna(0) 

GVC['Concept_name'] = GVC['Concept_name'].astype(str) 

GVC['Concept_Description'] = GVC['Concept_Description'].astype(str) 

GVC['Tags']=GVC['Concept_name']+GVC['Concept_Description'] 

print(GVC['Tags']) 

new_df=GVC[['S_no','Concept_name','Concept_Description','Tags',]] 

new_df.dropna(inplace=True) 

# new_df.isnull().sum() 

new_df['Tags']=new_df['Tags'].apply(lambda x:x.lower()) 

# print(new_df) 

import nltk 

from nltk.stem.porter import PorterStemmer 

ps = PorterStemmer() 

  

def stem(text): 

    y=[] 

    for i in text.split(): 

        y.append(ps.stem(i)) 

    return" ".join(y) 

  

  

new_df['Tags']=new_df['Tags'].apply(stem) 

print(new_df['Tags']) 

# new_df.head() 

# new_df.info() 

  

from sklearn.feature_extraction.text import CountVectorizer 

cv = CountVectorizer(max_features=4000,stop_words='english') 

  

vectors=cv.fit_transform(new_df['Tags']).toarray() 

vectors[0] 

feature_names=cv.get_feature_names_out() 

for feature in feature_names: 

    print(feature) 

  

from sklearn.metrics.pairwise import cosine_similarity 

similarity=cosine_similarity(vectors) 

similarity[2] 

  

new_df.info() 

new_df.head() 

  

#to remove unnecesary spaces. 

new_df['Concept_name'] = new_df['Concept_name'].str.strip() 

# new_df[new_df['Concept_name']=='Financial Reporting Standards'].index[0] 

def recommend(SSC): 

    course_index = new_df[new_df['Concept_name'] == SSC].index[0] 

    distances = similarity[course_index] 

    course_list = sorted(list(enumerate(distances)), reverse=True, key=lambda x: x[1])[1:6] 

  

    recommended_concepts = [] 

    for i in course_list: 

        recommended_concepts.append(new_df.iloc[i[0]].S_no) 

     

    return recommended_concepts 

  

  

GVC['Recommendation']=GVC['Concept_name'].apply(recommend) 

GVC['Recommendation'].head() 

  

file_path = "D:\\Python_Folder\\Course_Reco_Output.xlsx" 

GVC['Recommendation'].to_excel(file_path) 

print("Data has been written to the Excel file successfully.") 

# create new excel/ create new column if excel is already exist 

#already create a coumn header name 

#open excel sheet to write the data 

# append the data which coming fromm line no 72 

#%% 
import os

new_directory = "D:\Python_Folder"
os.chdir(new_directory)
print("Current Working Directory: ", os.getcwd())

import pandas as pd 
import numpy as np 

GVC= pd.read_excel("D:\Python_Folder\Coursera working mulltiple file.xlsx",sheet_name="Law")
# GVC.info()
# GVC.isnull().sum()

GVC['Course Name'] = GVC['Course Name'].fillna(0)
GVC['About Course '] = GVC['About Course '].fillna(0)
GVC['New_concepts'] = GVC['New_concepts'].fillna(0)
GVC['Course Name'] = GVC['Course Name'].astype(str) 
GVC['About Course '] = GVC['About Course '].astype(str) 
GVC['Tags']=GVC['Course Name']+GVC['About Course ']
GVC['Tags'].info()
# print(GVC['Tags']) 

new_df=GVC[['Course Name','New_concepts','Tags']] 
# new_df.dropna(inplace=True) 
# new_df.isnull().sum()

new_df['Tags']=new_df['Tags'].apply(lambda x:x.lower()) 
# print(new_df) 

import nltk
from nltk.stem.porter import PorterStemmer 
ps = PorterStemmer() 

def stem(text): 
      y=[] 
      for i in text.split():
          y.append(ps.stem(i)) 
      return" ".join(y) 


new_df['Tags']=new_df['Tags'].apply(stem)
# print(new_df['Tags']) 
# new_df.head() 


from sklearn.feature_extraction.text import CountVectorizer 
cv = CountVectorizer(max_features=11000,stop_words='english') 
vectors=cv.fit_transform(new_df['Tags']).toarray() 
vectors[0] 
feature_names=cv.get_feature_names_out() 
for feature in feature_names: 

    print(feature) 

  
from sklearn.metrics.pairwise import cosine_similarity 
similarity=cosine_similarity(vectors) 
similarity[2] 

      

# new_df.info() 

# new_df.head() 

#to remove unnecesary spaces. 

# new_df['Course Name'] = new_df['Course Name'].str.strip()

# new_df[new_df['Course Name']==''].index[0]

def recommend(SSC): 

    course_index = new_df[new_df['Course Name'] == SSC].index[0] 
    distances = similarity[course_index] 
    course_list = sorted(list(enumerate(distances)), reverse=True, key=lambda x: x[1])[1:6]
    recommended_concepts = []
    for i in course_list:
        recommended_concepts.append(new_df.iloc[i[0]].New_concepts) 
    return recommended_concepts 

GVC['Recommendation']=GVC['Course Name'].apply(recommend) 
GVC['Recommendation'].head() 
file_path = "D:\\Python_Folder\\Course_Reco_Output.xlsx"
GVC['Recommendation'].to_excel(file_path)
print("Data has been written to the Excel file successfully.")

# # create new excel/ create new column if excel is already exist 

# #already create a coumn header name 

# #open excel sheet to write the data 

# # append the data which coming fromm line no 72 
