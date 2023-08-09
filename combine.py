import pandas as pd
import os
import datetime
import sys
from sys import exit
from datetime import timedelta



def delete_files(directory = os.getcwd()): 
    """
    Deletes all the .xlsx files modified in the last one hour within the specified directory.

    This will always default to the current directory unless a parameter is specified.

    Parameters:
    directory (string, Optional): Directory where files should be deleted

    Returns:
    None

    """
    
    files = os.listdir(directory)
    for file in files:
        if file.endswith(".xlsx"):
            timestamp = os.path.getmtime(file) # File modification check
            datestamp = datetime.datetime.fromtimestamp(timestamp) # Convert timestamp into DateTime object
            current_time = datetime.datetime.now() # Using now() to get current time
            time_difference = current_time - datestamp # Get time delta
            if (time_difference.total_seconds() <= 3600): # 3600 seconds in one hour
                os.remove(os.path.join(directory, file)) # Delete files



def sheet_exists(excel_filename, sheetname): 
    """
    Returns TRUE if a sheet exists and FALSE if not.

    Returns a boolean value for if a sheet with a certain name exists in a given xlsx file.

    Parameters:
    excel_filename (string): The excel/xlsx filename without ".xlsx"
    sheetname (string): The name of the sheet that is being verified
    
    Returns:
    bool: True if the sheet exists, False if not

    """
    
    path = excel_filename + ".xlsx"
    try:
        sheets = pd.ExcelFile(path).sheet_names
        if sheetname in sheets:
            return True
        return False
    except:
        return False



def xlsx(dataframe, excel_filename, sheetname): 
    """
    Exports a Pandas DataFrame to an existing or new .xlsx file.

    Either makes both a new .xlsx file and a new sheet for the dataframe to go into, or only makes a new sheet into an existing .xlsx and adds the dataframe there.

    Parameters:
    dataframe (DataFrame): a single Pandas DataFrame which should be prepared to be exported to excel
    excel_filename (string): the excel/xlsx filename without ".xlsx"
    sheetname (string): the name of the sheet in the excel file that will contain the DataFrame
    
    Returns:
    None

    """
     
    path = excel_filename + ".xlsx"
    try: # First, try to add a new sheet to the current xlsx file if it already exists
        with pd.ExcelWriter(path, mode="a") as writer: 
            dataframe.to_excel(writer, sheet_name=sheetname)
    except: # If no xlsx file has yet been created, then make a new one
        dataframe.to_excel(path, sheet_name = sheetname)
    
    
    
def count_column_entries(df):
    return df.count(axis="columns")  
    
    
def get_files(directory = os.getcwd(), general_name = "Scribe", general_term = "Analysis"): # Default: Scribe_(number)_ADC_Analysis
    count = 0 # Keeps track of number of files found
    files = os.listdir(directory)
    relevant_filenames = []
    for file in files:
        if file.endswith(".xlsx"):
            general_name_location = file.find(general_name)
            general_term_location = file.find(general_term)
            if (general_name_location == 0 and general_term_location > 0):
                count = count+1
                relevant_filenames.append(file)
                
    return relevant_filenames



def merge_files(cleanup = True): # cleanup = True means remove all the random 0's
    file_list = get_files() # relevant .xlsx files list
    df_total = pd.DataFrame()
    df_list = []
    
    for file in file_list:
        excel_file = pd.ExcelFile(file)
        sheets = excel_file.sheet_names
        for sheet in sheets:
            if (sheet == "Gain"):
                df = excel_file.parse(sheet_name=sheet)
                df_list.append(df)
        
    #combo = [df1, df2]
    #result = pd.concat(combo)
    #print(result)
    #result.to_excel('combined_file.xlsx')
    count = 0
    for df in df_list:
        count = count + len(df)
    
    print(count)
    #result = pd.concat(df_list)
    #print(result)
    #result.to_excel('combined_file_GAIN.xlsx')
    
    #TODO need to deal with the random columns on the right of "Tad"... either remove it or have them next to tad in the space where each respective file actually starts in the final merged
    
    print("Done")

merge_files()




#https://pythoninoffice.com/use-python-to-combine-multiple-excel-files/
#https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.drop.html
#https://www.w3schools.com/python/pandas/ref_df_count.asp#:~:text=The%20count()%20method%20counts,each%20row%20(or%20column).
