__author__ = 'Chris'

import sys

sys.path.insert(0, 'Z:\Python projects\FishbowlAPITestProject')

import connecttest

def run_queries():
    myresults = connecttest.create_connection('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\SQL', 'BOMQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, 'Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'BOMs.xlsx')
    myresults = connecttest.create_connection('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\SQL', 'PartQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, 'Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'Parts.xlsx')
    myresults = connecttest.create_connection('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\SQL', 'MOQueryRedoux.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, 'Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'MOs.xlsx')
    myresults = connecttest.create_connection('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\SQL', 'POQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, 'Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'POs.xlsx')
    myresults = connecttest.create_connection('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\SQL', 'SOQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, 'Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'SOs.xlsx')
    myresults = connecttest.create_connection('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\SQL', 'INVQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, 'Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'INVs.xlsx')
    myresults = connecttest.create_connection('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\SQL', 'DescQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, 'Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'Descs.xlsx')
    return 'Queries Successful'

