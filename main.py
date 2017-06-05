import os
import sys
import glob
import re
import datetime
import shutil
import logging

import pandas as pd
from pandas import ExcelWriter

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

def main():
    logger.info("\n")

    csvNames = getCSVfiles(source_directory)

    sr = pd.Series(csvNames)
    grouped = sr.groupby(lambda x: re.search(r'\d+', sr[x]).group())

    for group in grouped:
        excelName = "File" + re.search(r'\d+', list(group[1])[0]).group() + ".xls"
        print (excelName)
        writer = ExcelWriter(excelName)
        for a_csv in group[1]:
            sheetName = re.search(r'[A-Za-z_\s]+', a_csv).group()[:-1]
            print sheetName
            try:
                df = pd.read_csv(os.path.join(source_directory, a_csv))
            except:
                df = pd.DataFrame()

            df.to_excel(writer,sheetName)
            writer.save()

    move_files(csvNames, source_directory, destination_directory)

if __name__ == "__main__":
    source_directory = source_directory

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    ch = logging.FileHandler(os.path.join(os.path.dirname(__file__), "task.log"))
    ch.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    today_string = datetime.datetime.today().strftime('%Y-%m-%d')
    destination_directory = os.path.join(source_directory, os.path.join("reports", today_string))

    main()





