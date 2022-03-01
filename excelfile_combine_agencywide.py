#!/usr/bin/env python
# coding: utf-8

# In[1]:


# -*- coding: utf-8 -*-

'''
Title: Combine survey responses from dozesn of files (1 survey question per Excel file) into a single Excel File
Author: Jessica Stockham
Date: March 1, 2022
'''


import pandas as pd
import csv

## Write tables to Excel Workbook
def multiple_dfs(df_list, caption_list, sheets, file_name, spaces):
    print("Writing with ExcelWriter")
    print(mapdict[key])
    writer = pd.ExcelWriter(file_name,engine='xlsxwriter')   
    row = 1
    
    for dataframe, caption in zip(df_list, caption_list):
        dataframe.to_excel(writer,sheet_name=sheets,startrow=row , startcol=0, index=False)   
        row = row + len(dataframe.index) + spaces + 1
        
        # Get the xlsxwriter workbook and worksheet objects.
        workbook = writer.book
        worksheet = writer.sheets[sheets]

         # Add a bold format to use to highlight cells.
        bold = workbook.add_format({'bold': True})
        
        # Write the caption.
        caption_spot = "A" + str(row - len(dataframe.index) - spaces - 1) 
        worksheet.write(caption_spot, caption, bold)
        
        # Get the dimensions of the dataframe.
        (max_row, max_col) = dataframe.shape

        # Make the columns wider for clarity.
        worksheet.set_column(0, max_col - 1, 25)
        worksheet.set_column('B:E', None)
        
        # Format the decimal as a percent
        format1 = workbook.add_format({'num_format': '0%'})
        worksheet.set_column('B:E', None, format1)
        
    writer.save()


# In[2]:


## Prepare Raw Data
def domainLoad(df_list, caption_list, key, table_list):
    for table in table_list:
        print(key)
        print("Preparing Tables")
        csvfile = table + ".csv"


        # Load in the CSV file with csv.reader to create a table caption
        
        with open(csvfile, newline='') as f:
            reader = csv.reader(f)
            row1 = next(reader)  # gets the first line (question number)
            row2 = next(reader) # gets the second line (question text)

            # For each Table, create caption out of Question Number & Question Text
            qnum = row1[2]
            qtext = row2[2]
            caption = qnum + ": " + qtext

            # For each Table, add to the caption list for the Key/Domain
            caption_list.append(caption)
   
        # For each Table, load in as its own datframe (Starts on row 3). Name the dataframe with the 'Table' name
        # Filter to needed columns, format to decimal, rename the row.name column with the question number
        table = pd.read_csv(csvfile, skiprows=3, usecols=[0, 1, 2, 3, 4, 5], encoding='Windows-1252')
        table = table[["row.name", "Program Agencies", "Compliance and Enforcement Agencies", "Management Agencies", "TotalPerc"]]
        
        # Total Count
        rslt_df = table.loc[table['row.name'] == 'TotalN']
       
        table = table.rename(columns={"row.name": qnum})
   

        # For each Table, add its Dataframe List for each Key/Domain
        df_list.append(table)
    
    # For each Domain, generate a sheet in the Excel Workbook 'ODEP.xlsx'
    multiple_dfs(df_list, caption_list, key, 'type.xlsx', 1)


# In[5]:


priority = ['QB_1', 
            'QB_7', 
            'QE_2', 
            'QE_3', 
            'QC_1',
            'QC_2', 
            'QC_3',
            'QE_4', 
            'QE_5']

secondary = ['QB_2A',   
             'QB_2B',
             'QB_3A',
             'QB_3B',
             'QB_4A',
             'QB_4B',
             'QB_5A',
             'QB_5B',
              'QC_1',
              'QC_2',
              'QC_3',
              'QF_2',
              'QF_4',
              'QE_1',
             'QB_6A_A1',
             'QB_6A_A2',
             'QB_6A_A3',
             'QB_6A_A4',
             'QB_6A_A5',
             'QB_6B_A2', 
             'QB_6B_A3',
             'QB_6B_A4',
              'QF_1_A1',
               'QF_1_A2',
               'QF_1_A3',
               'QF_1_A4',
               'QF_1_A5',
               'QF_1_A6',
               'QF_1_A7',
               'QF_1_A8',
                'QF_3',
                'QF_5']

coverage = ['QB_2A',   
             'QB_2B',
             'QB_3A',
             'QB_3B',
             'QB_4A',
             'QB_4B',
             'QB_5A',
             'QB_5B',
             'QB_6A_A2']

uses = ['QB_1']

effective = ['QE_1',
             'QB_6A_A1',
             'QB_6A_A3',
             'QB_6A_A4',
             'QB_6A_A5']
               

balance = ['QB_7']

quality = [  'QB_6B_A2', 
             'QB_6B_A3',
             'QB_6B_A4']
            
equity = ['QB_6B_A1']

capacity = ['QC_1',
              'QC_2',
              'QC_3',
              'QC_4',
              'QF_2',
              'QF_4', 
              'QG_1',
              'QG_2', 
              'QG_3', 
              'QG_4', 
               # Team Capacity
           
              'QE_2', 
              'QE_3'] 
              
context = ['QD_1',
             'QD_2',
             'QE_4', # Also in training (political support-only)
             'QE_5'] # Also in training (political support-only)

training = [
                'QE_4',
                'QE_5', 
               'QF_1_A1',
               'QF_1_A2',
               'QF_1_A3',
               'QF_1_A4',
               'QF_1_A5',
               'QF_1_A6',
               'QF_1_A7',
               'QF_1_A8',
                'QF_2',
                'QF_3',
                'QF_4',
                'QF_5']

# Respondent info: Sub-agency, GS-Level, Role/Responsibility
profile = ['QA_1',
           #'QA_2', 
           'QA_3', 
           'QA_4']

# Slides
slides = priority + secondary

# Dictionary MapDict = Domain (Key): Tables (Values)
mapdict = {
    "slides": slides
# "priority": priority,
# "secondary": secondary
#"coverage": coverage 
#'uses': uses
#"balance": balance
#"effective": effective
#"equity": equity,
#"capacity": capacity, 
#"context": context
#'quality': quality
#'training': training
#'profile': profile
}


# Domain Specific Loop (Loop through Keys/Domains)
for key in mapdict.keys():
    
    # Create list of Tables for each Domain
    table_list = mapdict[key]
    #print(table_list)
    
    # Create a List of Dataframes for each Key/Domain
    df_list = []

    # Create a List of Captions for each Key/Domain
    caption_list = []
   
    print(df_list, caption_list, key, table_list)
    
    # Table-Specific Loop (Loop through Tables in that Domain)
    domainLoad(df_list, caption_list, key, table_list)
    


# In[ ]:





# In[198]:





# In[46]:





# In[ ]:




