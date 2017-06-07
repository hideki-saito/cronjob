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
    extension = 'csv'
    os.chdir(directory)
    result = [i for i in glob.glob('*.{}'.format(extension))]

    return result

def move_files(fileList, source_directory, destination_directory):
    if not os.path.exists(destination_directory):
        os.makedirs(destination_directory)
    for file in fileList:
        shutil.move(os.path.join(source_directory, file), os.path.join(destination_directory, file))

def similiar_in_list(item, list):
    for i in list:
        if item.replace(" ", "").strip() == i.replace(" ", "").strip():
            return i

    return item

def main():
    logger.info("\n")

    csvNames = getCSVfiles(source_directory)

    sr = pd.Series(csvNames)
    grouped = sr.groupby(lambda x: re.search(r'\d+', sr[x]).group())

    for group in grouped:
        excelpath = os.path.join(source_directory, "File" + re.search(r'\d+', list(group[1])[0]).group() + ".xlsx")

        copyfile(os.path.join(current_directory, template), excelpath)

        book = load_workbook(excelpath)
        writer = ExcelWriter(excelpath, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        for a_csv in group[1]:
            csvName = re.search(r'[A-Za-z_\s]+', a_csv).group()[:-1]
            excel_df = pd.read_excel(excelpath, sheetname=csv_sheet[csvName])
            empty_excel_df = pd.DataFrame(columns=excel_df.columns)
            excel_columns = list(excel_df.columns)

            try:
                csv_df = pd.read_csv(os.path.join(source_directory, a_csv))
                csv_columns = pd.Series(list(csv_df.columns)).apply(similiar_in_list, list=excel_columns)
                csv_df.columns = csv_columns
            except:
                csv_df = pd.DataFrame()

            df = pd.concat([empty_excel_df, csv_df], ignore_index=True)[excel_columns]

            df.to_excel(writer, csv_sheet[csvName], startrow=len(excel_df)+1, header=False, columns=excel_columns, index=False)
            writer.save()

    move_files(csvNames, source_directory, destination_directory)

if __name__ == "__main__":
    source_directory = source_directory
    current_directory = os.path.dirname(__file__)
    template = template_excel
    today_string = datetime.datetime.today().strftime('%Y-%m-%d')
    destination_directory = os.path.join(source_directory, os.path.join("reports", today_string))

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    ch = logging.FileHandler(os.path.join(current_directory, "task.log"))
    ch.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    main()





