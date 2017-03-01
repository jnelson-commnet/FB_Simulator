__author__ = 'Chris'

import sys

sys.path.insert(0, "/mnt/manufacturing/Shared Services/Python projects/FishbowlAPITestProject")

import connecttest

def run_queries():
    myresults = connecttest.create_connection('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/SQL', 'BOMQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, '/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'BOMs.xlsx')
    myresults = connecttest.create_connection('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/SQL', 'PartQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, '/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'Parts.xlsx')
    myresults = connecttest.create_connection('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/SQL', 'MOQueryRedoux.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, '/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'MOs.xlsx')
    myresults = connecttest.create_connection('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/SQL', 'POQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, '/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'POs.xlsx')
    myresults = connecttest.create_connection('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/SQL', 'SOQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, '/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'SOs.xlsx')
    myresults = connecttest.create_connection('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/SQL', 'INVQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, '/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'INVs.xlsx')
    myresults = connecttest.create_connection('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/SQL', 'DescQuery.txt')
    myexcel = connecttest.makeexcelsheet(myresults)
    connecttest.save_workbook(myexcel, '/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'Descs.xlsx')
    return 'Queries Successful'

