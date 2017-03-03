__author__ = 'Chris'

import sys #Module to access the computer

sys.path.insert(0, 'Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux') #Pull up the file with the forecast information

import ForecastMain #Import the actual forecast python file

ForecastMain.run_normal_forecast_tiers_v2(ignore_schedule_errors=True) #Run the actual forecast