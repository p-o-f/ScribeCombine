import pandas as pd
import os
import datetime
import sys
from sys import exit
from datetime import timedelta



"""
# Many of the provided function below are UNCALLED since they're from another script that works with .xlsx files: https://github.com/p-o-f/TestDataParsing
# They are here for anyone who wants to maintain this code in the future. They may be useful or maybe not.
# Please read the associated READ_THIS files if you want to better understand or edit this code.
"""



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



def get_files(directory = os.getcwd(), general_name = "Scribe", general_term = "Analysis"): 
    """
    Returns a list of all .xlsx files in a directory which are named in a specific way. The default format is Scribe_(number)_..._Analysis (ex: Scribe11_ADC_Analysis)

    This will always default to the current directory unless a parameter is specified.

    Parameters:
    directory (string, Optional): Directory where files should be retrieved from
    general_name (string, Optional): The term the .xlsx filename should generally start with
    general_term (string, Optional): Another term which should generally in the .xlsx filename, but should also be exclusive to that filename (IE don't use "ADC" because there are ADChistograms and ADCanalysis files etc.)
    
    Returns:
    list[]

    """
    
    files = os.listdir(directory)
    relevant_filenames = []
    for file in files:
        if file.endswith(".xlsx"):
            general_name_location = file.find(general_name)
            general_term_location = file.find(general_term)
            if (general_name_location == 0 and general_term_location > 0): # Check if the file name begins with general_name and has general_term in it
                relevant_filenames.append(file)
                
    return relevant_filenames



def merge_sheets(file_list, sheetname):
    """
    Merges all sheets of a given name together.

    Across many .xlsx files, merges sheets with the same name and a similiar format together and returns a Pandas DataFrame of the merged result.

    Parameters:
    file_list (list[]): A list of .xlsx files that should all contain a sheet with the same name to be merged together
    sheetname (string): The name of the sheet that should be merged across many .xlsx files
    
    Returns:
    DataFrame

    """
    
    df_sheet_list = []
    found = False

    for file in file_list:
        found = False
        excel_file = pd.ExcelFile(file)
        sheets = excel_file.sheet_names
        for sheet in sheets:
            if (sheet == sheetname):
                df = excel_file.parse(sheet_name=sheet)
                df_sheet_list.append(df)
                found = True
        if (found is False): # This means the .xlsx file did not contain any instance of the sheet name that was being searched for
            print("Warning: the sheet name " + sheetname + " was not found in " + file + ". Please manually go back into this file and rename the sheet to be consistent with the other sheets.")
            print("Exiting script...")
            exit()

    combined_sheets = pd.concat(df_sheet_list) # Merge sheets together
    return combined_sheets 



def drop_columns(df, non_null_qty=2):
    """
    For a provided dataframe, drop columns that contain less than a certain number of non-null items. The reason for this is some of the Scribe sheets have the test name as a standalone column;
    in a master .xlsx file, all of these standalone columns would go next to each other at the top and serve no purpose for being there.

    I forgot the specifics of how this function works. It may be helpful to look up Pandas docs to gain more insight into what is actually happening. It is probably possible to just use DataFrame.dropna
    with axis=1 to achieve the same result, similar to the below drop_rows() function: https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.dropna.html

    Parameters:
    df (DataFrame): The Pandas DataFrame from which columns should be dropped
    non_null_qty (int, Optional): The threshold for non-null items; if there is less than this many non-null items in a column, then the column will be dropped
    
    Returns:
    DataFrame

    """
    
    result = df
    non_null_counts = df.notnull().sum()
    columns_to_drop = non_null_counts[non_null_counts < non_null_qty].index
    result = df.drop(columns=columns_to_drop)
    return result



def drop_rows(df, non_null_qty=2):
    """
    For a provided dataframe, drop rows that contain less than a certain number of non-null items. The reason for this is some of the Scribe sheets have rows with nothing but a 0 as an entry in one of the columns.
    These rows are redundant.

    See these docs for more details: https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.dropna.html

    Parameters:
    df (DataFrame): The Pandas DataFrame from which rows should be dropped
    non_null_qty (int, Optional): The threshold for non-null items; if there is less than this many non-null items in a row, then the row will be dropped
    
    Returns:
    DataFrame

    """
    
    result = df
    result = df.dropna(thresh=non_null_qty)
    return result



def merge_files(sheetname, cleanup = True): 
    """
    Merges and filters sheets, returning a DataFrame that is ready to be exported to xlsx.

    Goes through all relevant .xlsx files and merges their sheets with the same names together, then filtering them by dropping rows and columns, and finally returning a DataFrame that represents one
    large, mastercopy excel sheet which is ready to be exported to a .xlsx file.

    Parameters:
    sheetname (string): The name of the sheet that should be merged across many .xlsx files and made into a single file that can be exported to .xlsx
    cleanup (bool, Optional): setting this to TRUE (default) will call drop_rows() to remove the redundant rows where there is only a single 0 as an entry in any of the columns
    
    Returns:
    DataFrame

    """
    
    file_list = get_files() # The relevant .xlsx files 
    
    merged = merge_sheets(file_list, sheetname)
    
    if (cleanup is True): # Filter rows
        merged = drop_rows(merged)
    
    merged = drop_columns(merged) # Filter columns

    return merged



if __name__ == "__main__":
    """Driver function."""
    
    print("This is a script to combine .xlsx files labeled: Scribe(number)_ADC_Analysis together (i.e. Scribe5_ADC_Analysis). All of these files will be output into a new file with the name you choose.")
    print("\nPlease ensure all files that will be merged have the following sheets in them: MC, Gain, Offset, DNLmn, DNLmx, INL in this specific typecasing (order does not matter).")
    print("\nPlease also ensure all the files that will be merged are in the same directory as this script.")
    print("\nAny extraneous sheets will be disregarded in the master file and should be dealt with manually (IE pivot tables sheets labeled Worksheet, etc.)")
    print("\nRunning this script will take a relatively long time; expect 5-10 minutes to finish. Wait till the terminal says DONE to open the created master excel file. The terminal will give updates as progress is made.")
    print("\nPlease read the associated READ_THIS files if you want to better understand or edit this code.")
    print("----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
    
    proceed = input("If you understand the above, type Y to continue or anything else to exit the script: ")
    if (proceed != "Y"):
        print("Exiting...")
        exit()
    
    while True: # Prevent exporting to a blank file
        xlsx_path = input("\nEnter the .xlsx filename to export to: ")
        if xlsx_path != "" and xlsx_path.isspace() is False:
            break
    
    print("Merging sheets... please be patient")
    
    mc = merge_files("MC")
    gain = merge_files("Gain")
    offset = merge_files("Offset")
    DNLmn = merge_files("DNLmn")
    DNLmx = merge_files("DNLmx")
    inl = merge_files("INL")
    
    print("Sheet merging finished")
    print("Starting the export to .xlsx, this will take up to 5-10 minutes to finish...")
    
    xlsx(mc, xlsx_path, "MC")
    print("The MC sheet is exported. Starting the Gain sheet.")

    xlsx(gain, xlsx_path, "Gain")
    print("The Gain sheet is exported. Starting the Offset sheet.")

    xlsx(offset, xlsx_path, "Offset")
    print("The Offset sheet is exported. Starting the DNLmn sheet.")

    xlsx(DNLmn, xlsx_path, "DNLmn")
    print("The DNLmn sheet is exported. Starting the DNLmx sheet.")

    xlsx(DNLmx, xlsx_path, "DNLmx")
    print("The DNLmx sheet is exported. Starting the INL sheet.")
    
    xlsx(inl, xlsx_path, "INL")
    print("The INL sheet is exported.")
    print("DONE")