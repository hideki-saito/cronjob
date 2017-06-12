import os
import glob
import re
import datetime
import shutil
import logging
from shutil import copyfile

import pandas as pd
from pandas import ExcelWriter
from openpyxl import load_workbook
from config import *


def getCSVfiles(directory):
    '''
    Getting all csv file of the directory

    :param directory: target directory path
    :return: The path of all csv files in the target directory
    '''

    extension = 'csv'
    os.chdir(directory)
    result = [i for i in glob.glob('*.{}'.format(extension))]

    return result


def move_files(fileList, source_directory, destination_directory):
    '''
    Moving files from source directory to destination_directory

    :param fileList: List of files that will be moved
    :param source_directory: the path of the directory that contain the files now
    :param destination_directory: the path of the directory that will contain the files after execution the function.
    :return:
    '''

    # If the destination directory don't exist, it will make the directory. It avoids 'no such file or directory error'
    if not os.path.exists(destination_directory):
        os.makedirs(destination_directory)

    for file in fileList:
        shutil.move(os.path.join(source_directory, file), os.path.join(destination_directory, file))


def similiar_in_list(item, list):
    '''
    Find item from list that is same as a speicific string except spaces

    :param item: a speicific string
    :param list: the list of the strings
    :return: a string from the list that is simimliar to the argument string
    '''

    for i in list:
        if item.replace(" ", "").strip() == i.replace(" ", "").strip():
            return i

    return item

def main():
    '''
    Main function of the project. Create template, store info of csv into  sheets of the excel and move the csv files
    :return:
    '''

    logger.info("\n")

    csvNames = getCSVfiles(source_directory) # Get list of paths of all target csv files.

    sr = pd.Series(csvNames) # convert list to series of pandas
    groups = sr.groupby(lambda x: re.search(r'\d+', sr[x]).group())# grouping on above csv files by date

    # Looping on abvoe groups. Each group has csv files that each of these has name containing same date.
    for group in groups:
        # Define output excel name based on date that csv files of group have
        excelpath = os.path.join(source_directory, "File" + re.search(r'\d+', list(group[1])[0]).group() + ".xlsx")

        # If the excel file corresponding to the group does not exist, create a new excel file by template. If not, use existing the excel file.
        if not os.path.exists(excelpath):
            copyfile(os.path.join(current_directory, template), excelpath)

        # Preparing for adding rows to the new created excel file or the existing excel file.
        book = load_workbook(excelpath)
        writer = ExcelWriter(excelpath, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

        # Getting info from csv files and adding the info to the excel file.
        # Looping on csv files of the group
        for a_csv in group[1]:
            csvName = re.search(r'[A-Za-z_\s]+', a_csv).group()[:-1] #Get partial name of the csv file excepting date part.
            excel_df = pd.read_excel(excelpath, sheetname=csv_sheet[csvName]) # Get data of the excel sheet that corresponds to the above csvName
            empty_excel_df = pd.DataFrame(columns=excel_df.columns) # Create new empty pandas dataframe with same columns as the excel sheet
            excel_columns = list(excel_df.columns) # Get column names of the excel sheet

            try:
                csv_df = pd.read_csv(os.path.join(source_directory, a_csv), dtype=str) # Fetch info of the csv into pandas dataframe
                # Get the columns names of the csv file and adjust these to column names of the excel sheet.
                csv_columns = pd.Series(list(csv_df.columns)).apply(similiar_in_list, list=excel_columns)
                csv_df.columns = csv_columns # Alter the column names of the csv_df with adjusted column names, 'csv_columns'
            except:
                csv_df = pd.DataFrame()

            # Create new dataframe with same column names as the excel sheet by concatenateing emtpy_excel_df and csv_df and extracting data of excel_columns.
            df = pd.concat([empty_excel_df, csv_df], ignore_index=True)[excel_columns]

            # Add above df to the excel sheet starting next to existing rows.
            df.to_excel(writer, csv_sheet[csvName], startrow=len(excel_df)+1, header=False, columns=excel_columns, index=False)
            writer.save() # saving

    move_files(csvNames, source_directory, destination_directory) # Moving the processed csv files to the 'reports' folder.

if __name__ == "__main__":
    source_directory = source_directory # Get the path of the directory that contains target csv files from config.py
    current_directory = os.path.dirname(__file__) # Get the path of the directory
    template = template_excel # Get the path of the template excel file in config.py
    today_string = datetime.datetime.today().strftime('%Y-%m-%d') # Get today's date
    destination_directory = os.path.join(source_directory, os.path.join("reports", today_string)) # Get the path of the directory that the csv files will be moved to

    # Cutomizing of log mechanism
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    ch = logging.FileHandler(os.path.join(current_directory, "task.log"))
    ch.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    main() # run main function





