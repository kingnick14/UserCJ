from __future__ import print_function
import pandas as pd
import numpy as np
import glob
import os
import json
from argparse import ArgumentParser
from gooey import Gooey, GooeyParser

@Gooey(program_name="Combine SQL User Exports")

def parse_args():
    ###Use GooeyParser to build up the arguments we will use in our script
    ###Save the arguments in a default json file so that we can retrieve them
    ###every time we run the script.
    stored_args = {}
    # get the script name without the extension & use it to build up
    # the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    # Read in the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
    parser = GooeyParser(description='Combine SQL User Exports')
    parser.add_argument('data_directory',
                        action='store',
                        default=stored_args.get('data_directory'),
                        widget='DirChooser',
                        help="Source directory that contains Excel files")
    parser.add_argument('output_directory',
                        action='store',
                        widget='DirChooser',
                        default=stored_args.get('output_directory'),
                        help="Output directory to save summary report")
    # Do I need the below? perhaps as a "single client, or name chooser, or something??
    # parser.add_argument('cust_file',
    #                     action='store',
    #                     default=stored_args.get('cust_file'),
    #                     widget='FileChooser',
    #                     help='Customer Account Status File')

    args = parser.parse_args()
    # Store the values of the arguments so we have them next time we run
    with open(args_file, 'w') as data_file:
        # Using vars(args) returns the data as a dictionary
        json.dump(vars(args), data_file)
    return args



def exportcleaner(dtf,fname):
    dtf = dtf.loc[:, ~dtf.columns.str.contains('^Unnamed')]
    nextRow = 1
    while len(dtf.columns) == 1: #while the number of columns in the dataframe is 1, it means we read it wrong
        dtf = pd.read_excel(fname, header= nextRow) #use the next row as the header
        dtf = dtf.loc[:, ~dtf.columns.str.contains('^Unnamed')] #remove unnamed columns
        nextRow +=1 #add one to the counter
    return dtf

def combine_files(src_directory):
    all_data = pd.DataFrame()
    #print(glob.glob(src_directory,"*.xls*"))
    for f in glob.glob(os.path.join(src_directory,"*.xls*")):
        df = pd.read_excel(f)
        df = exportcleaner(df,f)
        #print(df)
        all_data = all_data.append(df,ignore_index=True,sort=True)
        print(f,"appended to the dataframe.")
    return all_data

#print(all_data.describe())
def save_results(combdata,output):
    ###Perform a summary of the data and save the data as an excel file###
    output_file = os.path.join(output, "combinedUserSQLs.xlsx")
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    combdata = combdata.reset_index()
    combdata.to_excel(writer)
    #all_data.to_excel(writer)
    #writer.save()

if __name__ == '__main__':
    conf = parse_args()
    print("Reading Excel files")
    user_df = combine_files(conf.data_directory)
    print("Saving sales and customer summary data")
    save_results(user_df, conf.output_directory)
    print("Done")
