# -*- coding: utf-8 -*-
# <nbformat>3.0</nbformat>

# <codecell>

import os
from sys import argv

script, source_file_input, save_file_input = argv

def csv_profile(source_file , save_file, working_directory = os.getcwd()):
    #Import 
    import openpyxl
    from datetime import datetime
    from pandas import ExcelWriter
    import pandas as pd
    import numpy as np
    import matplotlib.pyplot as plt
    from pandas import ExcelWriter
    pd.set_option('max_columns', 50)

    #Define functions
    def data_type_test(in_data):
        if type(in_data) == str:
            in_data = pd.to_datetime(in_data)
        return type(in_data)

    verbose = False
    save = True

    #Change current directory to documents folder
    os.chdir(working_directory)

    #Load in the csv file
    from_csv = pd.read_csv(source_file)

    print("Profiling " + working_directory + "/" + source_file)

    #Create table and populate with the data types
    data_profile = pd.DataFrame(from_csv.dtypes, columns=['data_type'])

    #Add profiling columns
    for i in data_profile.index.values:
        #Add Row count and number of unique values
        data_profile.loc[i,'Row_Count'] = len(from_csv[i])
        data_profile.loc[i,'Unique_Values'] = len(from_csv[i].value_counts())
        #Add first, second and third most common values
        data_profile.loc[i,'First_Most_Common'] = from_csv[i].value_counts().index[0]
        data_profile.loc[i,'First_Most_Common Count'] = from_csv[i].value_counts().values[0]
        if len(from_csv[i].value_counts()) > 1:
            data_profile.loc[i,'Second_Most_Common'] = from_csv[i].value_counts().index[1]
            data_profile.loc[i,'Second_Most_Common Count'] = from_csv[i].value_counts().values[1]
        if len(from_csv[i].value_counts()) > 2:    
            data_profile.loc[i,'Third_Most_Common'] = from_csv[i].value_counts().index[2]
            data_profile.loc[i,'Third_Most_Common Count'] = from_csv[i].value_counts().values[2]

    #Create summary stats for each column (
    summary_stats = from_csv.describe()
    summary_stats = pd.DataFrame.transpose(summary_stats)

    #Merge the data types and the stats together
    data_profile = pd.merge(data_profile, summary_stats, left_index=True, right_index=True, how='left')

    #Add additional information for the non numeric columns
    for i in data_profile.index.values:
        if data_profile.loc[i,'data_type'] == "object":
            data_profile.loc[i,'data_type'] = str(data_type_test(from_csv.loc[1,i]).__name__)
            data_profile.loc[i,'count'] = from_csv[i].count()
            if data_profile.loc[i,'data_type'] == 'Timestamp':
                data_profile.loc[i,'max'] = max(pd.to_datetime(from_csv[i]))
                data_profile.loc[i,'min'] = min(pd.to_datetime(from_csv[i]))

    #The data type column needs to be converted to stings before being sent to the excel file
    data_profile.data_type = data_profile.data_type.apply(str)

    if save:
        writer = ExcelWriter(save_file)
        data_profile.to_excel(writer, 'sheet1')
        writer.save()
        print("Profile saved to " + working_directory + "/" + save_file)

print("Running ",script)
csv_profile(source_file_input, save_file_input)

