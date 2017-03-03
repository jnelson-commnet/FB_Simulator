__author__ = 'Chris'

import os
import sys

sys.path.insert(0, 'Z:\Python projects\FishbowlAPITestProject')

import connecttest

print('API Paths')
homey = os.getcwd()
redouxPath = os.path.join(homey, 'ForecastRedoux')
sqlPath = os.path.join(redouxPath, 'SQL')
rawDataPath = os.path.join(redouxPath, 'RawData')
print(sqlPath)
print(rawDataPath)

def run_queries():
    myresults = connecttest.create_connection(sqlPath, 'BOMQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, rawDataPath, 'BOMs.xlsx')
    myresults = connecttest.create_connection(sqlPath, 'PartQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, rawDataPath, 'Parts.xlsx')
    myresults = connecttest.create_connection(sqlPath, 'MOQueryRedoux.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, rawDataPath, 'MOs.xlsx')
    myresults = connecttest.create_connection(sqlPath, 'POQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, rawDataPath, 'POs.xlsx')
    myresults = connecttest.create_connection(sqlPath, 'SOQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, rawDataPath, 'SOs.xlsx')
    myresults = connecttest.create_connection(sqlPath, 'INVQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, rawDataPath, 'INVs.xlsx')
    myresults = connecttest.create_connection(sqlPath, 'DescQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, rawDataPath, 'Descs.xlsx')
    return 'Queries Successful'

