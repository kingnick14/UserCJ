import pandas as pd
import numpy as np
import os
import sys
import glob

def exportcleaner(dtf,fname):
    dtf = dtf.loc[:, ~dtf.columns.str.contains('^Unnamed')]
    nextRow = 1
    while len(dtf.columns) == 1: #while the number of columns in the dataframe is 1, it means we read it wrong
        dtf = pd.read_excel(fname, header= nextRow) #use the next row as the header
        dtf = dtf.loc[:, ~dtf.columns.str.contains('^Unnamed')] #remove unnamed columns
        nextRow +=1 #add one to the counter
    return dtf
#target = input('Please specify the directory')
# #print(glob.glob("../pandas/sales*.xlsx"))
#

# dirlist = os.listdir(target)
#     #target = open(target, "r+")
all_data = pd.DataFrame()
print(glob.glob("../UserCJ/data/*.xls*"))
for f in glob.glob("../UserCJ/data/*.xls*"):
    df = pd.read_excel(f)
    df = exportcleaner(df,f)
    #print(df)
    all_data = all_data.append(df,ignore_index=True,sort=True)
    print(f,"appended to the dataframe.")

print(all_data.describe())

writer = pd.ExcelWriter('output.xlsx')
all_data.to_excel(writer)
writer.save()
