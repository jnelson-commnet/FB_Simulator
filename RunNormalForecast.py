__author__ = 'Chris'

import sys #Module to access the computer

sys.path.insert(0, 'Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux') #Pull up the file with the forecast information

import ForecastMain #Import the actual forecast python file

ForecastMain.run_normal_forecast() #Run the actual forecast