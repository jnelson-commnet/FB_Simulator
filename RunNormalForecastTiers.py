__author__ = 'Chris'


import sys #Module to access the computer
import os

homey = os.getcwd()
forcRedouxPath = os.path.join(homey, 'ForecastRedoux')


sys.path.insert(0, forcRedouxPath) #Pull up the file with the forecast information

import ForecastMain #Import the actual forecast python file

ForecastMain.run_normal_forecast_tiers_v2() #Run the actual forecast