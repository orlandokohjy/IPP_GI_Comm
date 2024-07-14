# import the required packages
from datetime import date
import datetime
import pandas as pd
import numpy as np
import os
import glob
import re

import nltk
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer

from fuzzywuzzy import fuzz
from fuzzywuzzy import process


class Name_matching:
    def __init__(self, folder_path, all_advisor_df):
        self.folder_path = folder_path
        self.all_advisor_df = all_advisor_df

    def preprocess_text(self, text):
        # convert to lowercase
        text = text.lower()
        
        # remove punctuation
        text = re.sub(r'[^\w\s]', '', text)
        
        # remove numbers
        text = re.sub(r'\d+', '', text)
        
        # remove stopwords
        stop_words = set(stopwords.words('english'))
        words = text.split()
        words = [word for word in words if word not in stop_words]
        text = ' '.join(words)
        
        
        # lemmatize
        lemmatizer = WordNetLemmatizer()
        words = text.split()
        words = [lemmatizer.lemmatize(word) for word in words]
        text = ' '.join(words)
        
        return text



    def matching_far(self):
        # load the Master Record list for the referrer list
        file_pattern = 'W4*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        master_record = pd.read_excel(matching_files[0])

        # load the e-submission list for the referrer list
        file_pattern = 'W3*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        e_sub = pd.read_excel(matching_files[0])



        # remove the N.A in the mixed column of "Tfar" in the master_record
        master_tfar = master_record[master_record['Tfar'].str.contains('N.A|N.A.|YES', na=False)]['Tfar']

        # remove the date in the mixed column of "Referer" in the master_record
        master_record['Referer_mix'] = pd.to_datetime(master_record['Referer'], errors='coerce')

        # filter out the datetime values
        master_ref = master_record[(master_record['Referer_mix'].isna()) & (~master_record['Referer'].str.contains('0|/', na=False)) & (master_record['Referer'].notna())]['Referer']

        # concatenate the two series
        master_adviser = pd.concat([master_tfar, master_ref])
        
        
        # remove the Not Applicable
        e_sub_ref = e_sub[e_sub['Name of Referral'].str.contains('Not', na=False)]['Name of Referral']

        # concatenate TFAR and Referrer series
        e_sub_adviser = pd.concat([e_sub['Name of TFAR'], e_sub_ref])

        e_sub_adviser = pd.DataFrame(e_sub_adviser).rename(columns={0:'ADVISER'})

        # get all the adviser names here
        comm_far = pd.concat([self.all_advisor_df['TFAR'], e_sub_adviser['ADVISER']], axis=0).dropna().drop_duplicates().astype(str)

        # convert the above series to a DataFrame
        comm_far = pd.DataFrame(comm_far, columns=['ADVISER'])

        # load the FAR masterlist
        file_pattern = 'W1*.xls*' 

        matching_files = glob.glob(os.path.join(self.folder_path, file_pattern))

        master_list = pd.read_excel(matching_files[0], skiprows=1)

        # first row number of the row that has all the values to be NaN. 
        first_row_with_all_NaN = master_list[master_list.isnull().all(axis=1) == True].index.tolist()[0]

        # get the FAR complete name from the master list
        master_list = master_list.iloc[0:first_row_with_all_NaN-1]
        master_list = master_list['NEW FAR NAME']

        # text preprocessing for comm_far
        comm_far['Processed_Name'] = comm_far['ADVISER'].apply(lambda x: self.preprocess_text(x))


        # empty lists for storing the matches later
        match_1 = []
        match_2 = []
        p = []

        # converting dataframe column to list of elements to do fuzzy matching
        comm_far_list = comm_far['Processed_Name']#.tolist()


        # taking the threshold as 80
        threshold = 80

        # iterating through comm_far_list to extract its closest match from the master_list
        for i in comm_far_list:
            match_1.append(process.extractOne(i, master_list.astype(str), scorer=fuzz.ratio))
            
        comm_far['matched_name'] = list(zip(*match_1))[0]
        comm_far['matching_score'] = list(zip(*match_1))[1]
        
        return comm_far