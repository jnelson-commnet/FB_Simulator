__author__ = 'Chris'

import ForecastAPI
import os
import pandas as pd
import ForecastTimelineBackend as ftlb
import ForecastSettings as fs
import xlsxwriter
import datetime as dt

"""Pull the orders data"""
def gather_orders():
    mopath = os.path.join('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'MOs.xlsx')
    modf = pd.read_excel(mopath, header=0) #Opens and puts the data into a dataframe
    orgdf = ftlb.create_mo_parents(modf)
    popath = os.path.join('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'POs.xlsx')
    tempdf = pd.read_excel(popath, header=0) #Opens and puts the data into a dataframe
    orgdf = orgdf.append(tempdf.copy()) #Append all the data into one dataframe
    sopath = os.path.join('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'SOs.xlsx')
    tempdf = pd.read_excel(sopath, header=0) #Opens and puts the data into a dataframe
    orgdf = orgdf.append(tempdf.copy()) #Append all the data into one dataframe
    orgdf = orgdf.sort_values(by=['PART', 'DATESCHEDULED'], ascending=[True, False]) #Sort the data by part then date
    return orgdf

"""Pull the parts data"""
def gather_parts():
    partspath = os.path.join('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'Parts.xlsx')
    partsdf = pd.read_excel(partspath) #Opens and puts the data into a dataframe
    partsdf = partsdf.sort_values(by='PART', ascending=True) #Sort the data by part
    return partsdf

"""Pull the inventory data"""
def gather_inventory():
    invpath = os.path.join('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'INVs.xlsx')
    orgdf = pd.read_excel(invpath, header=0) #Opens and puts the data into a dataframe
    orgdf = orgdf.sort_values(by='PART', ascending=True) #Sort the data by part
    return orgdf

"""Pull the BOM data"""
def gather_boms():
    bompath = os.path.join('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData', 'BOMs.xlsx')
    orgdf = pd.read_excel(bompath, header=0) #Opens and puts the data into a dataframe
    orgdf = orgdf.sort_values(by=['BOM','FG','PART'], ascending=[True,True,True]) #Sort the data by bom, finished good, then part
    return orgdf

"""Create make/buy table."""
def make_buy(mbdf):
    mbdf = mbdf.drop('AvgCost', axis=1) #Drop the avg cost column. If we want to use this later we can
    mbdf = mbdf.sort_values(by='PART', ascending=True) #Sort the data by part
    mbdf = mbdf.drop_duplicates('PART') #Don't think this is actually neccesary
    return mbdf

"""Add make/buy column to main table."""
def add_make_buy(orgdf, mbdf):
    newdf = pd.merge(orgdf, mbdf, on='PART', how='left') #Add on the make buy column to the regular orders dataframe
    return newdf

"""Adds new orders to original orders"""
def add_new_orders(ordersdf, newordersdf):
    allordersdf = ordersdf.append(newordersdf) #Add the new orders created from shortages
    allordersdf.reset_index(drop=True, inplace=True)
    # print(allordersdf)
    return allordersdf.copy()

"""Just combining all the data pulling functions. This isn't used.
    It's just a quick function to combine all the data gathering functions."""
def data_prep():
    ordersdf = gather_orders()
    partsdf = gather_parts()
    mbdf = make_buy(partsdf)
    newordersdf = add_make_buy(ordersdf, mbdf)
    invdf = gather_inventory()
    bomsdf = gather_boms()
    return [newordersdf, invdf, bomsdf]

"""
To avoid extra orders due to scheduling issues, this can be run after data_prep().
It takes all the positive inventory changes and bumps them back a few years earlier (so PO's are counted first before usage).
"""
def bandaid_schedule_issues(ordersdf):
    past = ordersdf[ordersdf['QTYREMAINING'] > 0].copy()  # all positive inv changes
    future = ordersdf[ordersdf['QTYREMAINING'] < 0].copy()  # all negative inv changes
    today = dt.datetime.now()  # grab today's date
    yesteryear = today + dt.timedelta(days=-2000)  # bump it back a few years
    past['DATESCHEDULED'] = yesteryear.strftime('%Y-%m-%d %H:%M:%S')  # copy it into the positive inv changes
    ordersdf = past.copy().append(future.copy())  # append all inv changes back together (will have dropped anything equal to 0)
    ordersdf.sort_values(by=['PART', 'DATESCHEDULED'], ascending=[True, False], inplace=True)  # sort it
    return ordersdf



"""Loops to capture all orders"""
def complete_orders_loop_redoux(missingboms, manyboms):
    ordersdf = gather_orders()
    partsdf = gather_parts()
    mbdf = make_buy(partsdf)
    newordersdf = add_make_buy(ordersdf, mbdf)
    invdf = gather_inventory()
    bomsdf = gather_boms()
    column_headers = ['ORDER', 'ITEM', 'ORDERTYPE', 'PART', 'QTYREMAINING', 'DATESCHEDULED', 'PARENT', 'Make/Buy', 'PriorInv'] #Create Headers list for future use
    finalordersdf = pd.DataFrame(columns=column_headers) #Create dataframe which will be the final product
    shortorderslist = ftlb.find_next_shortage_redoux(invdf, newordersdf) #This returns the list below and finds the shortages for every part
    # shortagedf = shortorderslist[0]
    # postordersdf = shortorderslist[1]
    # preordersdf = shortorderslist[2]
    # invdf = shortorderslist[3]
    finalordersdf = finalordersdf.append(shortorderslist[2]) #Append the orders covererd by inventory
    neworders = ftlb.make_new_orders(shortorderslist[0], bomsdf, missingboms, manyboms) #Make phantom orders to make up for the shortages
    shortorderslist[1] = add_new_orders(shortorderslist[1], neworders) #Append the phantom orders to the orders timeline
    while shortorderslist[1].empty == False: #Run this loop until there are no more orders left to parse through
        shortorderslist = ftlb.find_next_shortage_redoux(shortorderslist[3], shortorderslist[1]) #Finds shortages for every part
        finalordersdf = finalordersdf.append(shortorderslist[2]) #Appends the covered orders to the final product
        if shortorderslist[0].empty: #If there are no shortages break the loop
            break
        neworders = ftlb.make_new_orders(shortorderslist[0], bomsdf, missingboms, manyboms) #Make phantom orders from the shortages
        shortorderslist[1] = add_new_orders(shortorderslist[1], neworders) #Add the phantom orders to the orders timeline
    finalordersdf = finalordersdf.sort_values(by=['PART', 'DATESCHEDULED'], ascending=[True, True])
    finalordersdf.reset_index(drop=True, inplace=True) #Finally sort and reindex the final product to return
    return [finalordersdf, shortorderslist[3]]

"""Loops to capture all orders for different tiers"""
def complete_orders_loop_redux(partlist, ordersdf, invdf, bomsdf, missingboms, manyboms):
    column_headers = ['ORDER', 'ITEM', 'ORDERTYPE', 'PART', 'QTYREMAINING', 'DATESCHEDULED', 'PARENT', 'Make/Buy'] #Create Headers list for future use
    newordersdf = pd.DataFrame(columns=column_headers)  # Save an empty dataframe with pre set column headers, the orders to run will be saved here consecutively.
    ordersdf.reset_index(drop=True, inplace=True)
    # partlist = [x for x in partlist if str(x) != 'nan'] #Only for excel  ### I don't think this is necessary on a full run, just when saving state with excel - Jack
    for index, each in ordersdf.iterrows():  # Checking every row in ordersdf
        if str(each['PART']) in partlist:  # If the part number in ordersdf is also in the current tierlist of parts to run ...
            templist = [each['ORDER'], each['ITEM'], each['ORDERTYPE'], each['PART'], each['QTYREMAINING'], each['DATESCHEDULED'], each['PARENT'], each['Make/Buy']]  # then save that row ...
            newordersdf.loc[len(newordersdf)] = templist  # and add it to newordersdf.
            ordersdf.drop([index], inplace=True)  # drop it from the original orderlist to avoid revisiting it ### I'm not sure how this line is needed yet - Jack
            # print(newordersdf)
    finalordersdf = pd.DataFrame(columns=column_headers) #Create dataframe which will be the final product
    # print(newordersdf)
    shortorderslist = ftlb.find_next_shortage_redux(invdf, newordersdf, partlist) #This returns the list below and finds the shortages for every part
    # shortagedf = shortorderslist[0]
    # postordersdf = shortorderslist[1]
    # preordersdf = shortorderslist[2]
    # invdf = shortorderslist[3]
    finalordersdf = finalordersdf.copy().append(shortorderslist[2].copy()) #Append the orders covered by inventory
    neworders = ftlb.make_new_orders(shortorderslist[0].copy(), bomsdf.copy(), missingboms, manyboms) #Make phantom orders to make up for the shortages
    shortorderslist[1] = add_new_orders(shortorderslist[1].copy(), neworders.copy()) #Append the phantom orders to the orders timeline
    # print(shortorderslist[1])
    while shortorderslist[1].empty == False: #Run this loop until there are no more orders left to parse through
        # print(shortorderslist[0])
        shortorderslist = ftlb.find_next_shortage_redux(shortorderslist[3], shortorderslist[1], partlist) #Finds shortages for every part
        finalordersdf = finalordersdf.copy().append(shortorderslist[2].copy()) #Appends the covered orders to the final product
        if shortorderslist[0].empty: #If there are no shortages break the loop
            break
        neworders = ftlb.make_new_orders(shortorderslist[0].copy(), bomsdf.copy(), missingboms, manyboms) #Make phantom orders from the shortages
        shortorderslist[1] = add_new_orders(shortorderslist[1].copy(), neworders.copy()) #Add the phantom orders to the orders timeline
        # print(shortorderslist[1])
    # print(shortorderslist[1])
    ordersdf = ordersdf.copy().append(shortorderslist[1].copy()) # This S.O.B. took me a while to track down ... just added this so new orders pass to the next tier ... damn
    finalordersdf = finalordersdf.sort_values(by=['PART', 'DATESCHEDULED'], ascending=[True, True])
    finalordersdf.reset_index(drop=True, inplace=True) #Finally sort and reindex the final product to return
    # print(ordersdf)
    return [finalordersdf.copy(), shortorderslist[3].copy(), ordersdf.copy()]

"""Run forecast on fake order. This goes through the same process as complete_orders_loop_redoux.
    First it grabs the parts from parts to build and it creates those orders and adds them to the timeline."""
def run_fake_orders(missingboms, manyboms):
    newbuilds = pd.read_excel('Z:\Python projects\Chris Forecast\Forecast\AdditionalInfo\PartsToBuild.xlsx', sheetname='Sheet1')
    ordersdf = gather_orders()
    partsdf = gather_parts()
    mbdf = make_buy(partsdf)
    newordersdf = add_make_buy(ordersdf, mbdf)
    invdf = gather_inventory()
    bomsdf = gather_boms()
    for index, row in newbuilds.iterrows():
        bomnum = row['Part']
        bomqty = row['Qty']
        bomdate = row['Date']
        bomdate = dt.datetime.strptime(bomdate, '%Y-%m-%d')
        bomorderdf = ftlb.create_phantom_order(bomnum, bomqty, bomdate, bomsdf, missingboms, manyboms)
        newordersdf = ftlb.add_phantom_order(bomorderdf, newordersdf)
    column_headers = ['ORDER', 'ITEM', 'ORDERTYPE', 'PART', 'QTYREMAINING', 'DATESCHEDULED', 'PARENT', 'Make/Buy']
    finalordersdf = pd.DataFrame(columns=column_headers)
    shortorderslist = ftlb.find_next_shortage_redoux(invdf, newordersdf)
    # shortagedf = shortorderslist[0]
    # postordersdf = shortorderslist[1]
    # preordersdf = shortorderslist[2]
    # invdf = shortorderslist[3]
    finalordersdf = finalordersdf.append(shortorderslist[2])
    neworders = ftlb.make_new_orders(shortorderslist[0], bomsdf, missingboms, manyboms)
    shortorderslist[1] = add_new_orders(shortorderslist[1], neworders)
    while shortorderslist[1].empty == False:
        shortorderslist = ftlb.find_next_shortage_redoux(shortorderslist[3], shortorderslist[1])
        finalordersdf = finalordersdf.append(shortorderslist[2])
        if shortorderslist[0].empty:
            break
        neworders = ftlb.make_new_orders(shortorderslist[0], bomsdf, missingboms, manyboms)
        shortorderslist[1] = add_new_orders(shortorderslist[1], neworders)
    finalordersdf = finalordersdf.sort_values(by=['PART', 'DATESCHEDULED'], ascending=[True, True])
    finalordersdf.reset_index(drop=True, inplace=True)
    return [finalordersdf, shortorderslist[3]]

"""Takes away any orders that are known to not be included in the demand for a part.
    This goes through the same process as complete_orders_loop_redoux.
    First it grabs the parts from parts to remove and it removes those orders from the timeline."""
def remove_orders(missingboms, manyboms):
    droporders = pd.read_excel('Z:\Python projects\Chris Forecast\Forecast\AdditionalInfo\OrdersToRemove.xlsx', sheetname='Sheet1')
    ordersdf = gather_orders()
    partsdf = gather_parts()
    mbdf = make_buy(partsdf)
    newordersdf = add_make_buy(ordersdf, mbdf)
    invdf = gather_inventory()
    bomsdf = gather_boms()
    newtimeline = newordersdf
    for item, row in droporders.iterrows():
        if row['OrderType'] == 'MO':
            newtimeline = newtimeline.ix[newtimeline['ITEM'] != row['OrderNum']].copy()
        else:
            newtimeline = newtimeline.ix[newtimeline['ORDER'] != str(row['OrderNum'])].copy()
    column_headers = ['ORDER', 'ITEM', 'ORDERTYPE', 'PART', 'QTYREMAINING', 'DATESCHEDULED', 'PARENT', 'Make/Buy']
    finalordersdf = pd.DataFrame(columns=column_headers)
    shortorderslist = ftlb.find_next_shortage_redoux(invdf, newtimeline)
    # shortagedf = shortorderslist[0]
    # postordersdf = shortorderslist[1]
    # preordersdf = shortorderslist[2]
    # invdf = shortorderslist[3]
    finalordersdf = finalordersdf.append(shortorderslist[2])
    neworders = ftlb.make_new_orders(shortorderslist[0], bomsdf, missingboms, manyboms)
    shortorderslist[1] = add_new_orders(shortorderslist[1], neworders)
    while shortorderslist[1].empty == False:
        shortorderslist = ftlb.find_next_shortage_redoux(shortorderslist[3], shortorderslist[1])
        finalordersdf = finalordersdf.append(shortorderslist[2])
        if shortorderslist[0].empty:
            break
        neworders = ftlb.make_new_orders(shortorderslist[0], bomsdf, missingboms, manyboms)
        shortorderslist[1] = add_new_orders(shortorderslist[1], neworders)
    finalordersdf = finalordersdf.sort_values(by=['PART', 'DATESCHEDULED'], ascending=[True, True])
    finalordersdf.reset_index(drop=True, inplace=True)
    return [finalordersdf, shortorderslist[3]]

"""Run the normal forecast. This does all the outside work. Pull data from API create the excel sheet and save it."""
def run_normal_forecast():
    # sql = ForecastAPI.run_queries() #Runs the function in ForecastAPI that pulls data from Fishbowl
    # print(sql) #Prints Queries Successful!
    missingboms = fs.No_BOMs() #Creates an instance of the class in Forecast Settings No_BOMs
    manyboms = fs.Many_BOMs() #Creates an instance of the class in Forecast Settings Many_BOMs
    normal_orders = complete_orders_loop_redoux(missingboms, manyboms) #Runs the function that actually builds out the timeline
    print('*Timeline Has Been Created*')
    timingtest = ftlb.find_timing_issues(normal_orders[0], normal_orders[1]) #Finds parts that have phantom orders and are left with a positive inventory
    demand = ftlb.find_demand_driver(normal_orders[0]) #Runs the loop that figures out the top level driver for all the orders
    phantoms = ftlb.get_phantom_orders(demand) #Returns all the phantom orders
    writer = pd.ExcelWriter('Z:\Python projects\Chris Forecast\Forecast\mytimelinetest.xlsx') #Creates a test excel file
    demand.to_excel(writer, 'Sheet') #Fills the test excel with the whole timeline
    writer.save()
    workbook = xlsxwriter.Workbook('Z:\Python projects\Chris Forecast\Forecast\RegularForecast.xlsx') #Create the actual final product excel file
    orderslist = ftlb.split_phantoms(phantoms) #Seperates purchase and manufacture phantom orders
    worksheetP = workbook.add_worksheet('Purchasing') #Create the seperate worksheets for purchasing and manufacturing
    worksheetM = workbook.add_worksheet('Manufacturing')
    ftlb.create_subtotals_format(workbook=workbook, worksheet=worksheetP, thedf=orderslist[0], timinglist=timingtest) #This function creates the subtotals in the worksheet
    ftlb.create_subtotals_format(workbook=workbook, worksheet=worksheetM, thedf=orderslist[1], timinglist=timingtest) #Then is calls another function to add the data in FTLB
    worksheetT = workbook.add_worksheet('Timeline') #Adds another worksheet for the timeline
    ftlb.create_timeline_worksheet(workbook, worksheetT, demand) #Fills the timeline worksheet with data
    worksheetN = workbook.add_worksheet('PartsWithNoBOM') #Adds another worksheet for No BOMs
    noboms = missingboms.get_parts() #Returns the parts with no BOMs
    ftlb.create_miss_bom_worksheet(worksheetN, noboms) #Fills the worksheet with data
    worksheetL = workbook.add_worksheet('PartsWithTooManyBOMs') #Creates another worksheet for too many BOMs
    lotsboms = manyboms.get_parts() #Returns the parts with many active BOMs
    ftlb.create_too_many_boms_worksheet(worksheetL, lotsboms) #Fills the worksheet with data
    workbook.close() #Closing the workbook commits all the data
    print('*The forecast is done!*')

"""Run the normal forecast. This does all the outside work. Pull data from API create the excel sheet and save it."""
def run_normal_forecast_tiers():
    # sql = ForecastAPI.run_queries() #Runs the function in ForecastAPI that pulls data from Fishbowl
    # print(sql) #Prints Queries Successful!
    datalist = data_prep()
    normal_orders = [datalist[0], datalist[1]]
    invdf = datalist[1]
    bomsdf = datalist[2]
    mypartsdf = gather_parts()
    mytierlist = ftlb.create_bom_tiers(bomsdf, mypartsdf)
    missingboms = fs.No_BOMs() #Creates an instance of the class in Forecast Settings No_BOMs
    manyboms = fs.Many_BOMs() #Creates an instance of the class in Forecast Settings Many_BOMs
    for key, value in mytierlist.items():
        normal_orders = complete_orders_loop_redux(value, normal_orders[0], invdf, bomsdf, missingboms, manyboms) #Runs the function that actually builds out the timeline
    print('*Timeline Has Been Created*')
    timingtest = ftlb.find_timing_issues(normal_orders[0], normal_orders[1]) #Finds parts that have phantom orders and are left with a positive inventory
    demand = ftlb.find_demand_driver(normal_orders[0]) #Runs the loop that figures out the top level driver for all the orders
    phantoms = ftlb.get_phantom_orders(demand) #Returns all the phantom orders
    writer = pd.ExcelWriter('Z:\Python projects\Chris Forecast\Forecast\mytimelinetest.xlsx') #Creates a test excel file
    demand.to_excel(writer, 'Sheet') #Fills the test excel with the whole timeline
    writer.save()
    workbook = xlsxwriter.Workbook('Z:\Python projects\Chris Forecast\Forecast\RegularForecast.xlsx') #Create the actual final product excel file
    orderslist = ftlb.split_phantoms(phantoms) #Seperates purchase and manufacture phantom orders
    worksheetP = workbook.add_worksheet('Purchasing') #Create the seperate worksheets for purchasing and manufacturing
    worksheetM = workbook.add_worksheet('Manufacturing')
    ftlb.create_subtotals_format(workbook=workbook, worksheet=worksheetP, thedf=orderslist[0], timinglist=timingtest) #This function creates the subtotals in the worksheet
    ftlb.create_subtotals_format(workbook=workbook, worksheet=worksheetM, thedf=orderslist[1], timinglist=timingtest) #Then is calls another function to add the data in FTLB
    worksheetT = workbook.add_worksheet('Timeline') #Adds another worksheet for the timeline
    ftlb.create_timeline_worksheet(workbook, worksheetT, demand) #Fills the timeline worksheet with data
    worksheetN = workbook.add_worksheet('PartsWithNoBOM') #Adds another worksheet for No BOMs
    noboms = missingboms.get_parts() #Returns the parts with no BOMs
    ftlb.create_miss_bom_worksheet(worksheetN, noboms) #Fills the worksheet with data
    worksheetL = workbook.add_worksheet('PartsWithTooManyBOMs') #Creates another worksheet for too many BOMs
    lotsboms = manyboms.get_parts() #Returns the parts with many active BOMs
    ftlb.create_too_many_boms_worksheet(worksheetL, lotsboms) #Fills the worksheet with data
    workbook.close() #Closing the workbook commits all the data
    print('*The forecast is done!*')

"""Run forecast with an added order. This does not pull fresh data!
    This does the same thing as run_normal-forecast"""
def run_add_order_forecast():
    missingboms = fs.No_BOMs()
    manyboms = fs.Many_BOMs()
    add_order = run_fake_orders(missingboms, manyboms)
    timingtest = ftlb.find_timing_issues(add_order[0], add_order[1])
    demand = ftlb.find_demand_driver(add_order[0])
    phantoms = ftlb.get_phantom_orders(demand)
    workbook = xlsxwriter.Workbook('Z:\Python projects\Chris Forecast\Forecast\AddOrderForecast.xlsx')
    orderslist = ftlb.split_phantoms(phantoms)
    worksheetP = workbook.add_worksheet('Purchasing')
    worksheetM = workbook.add_worksheet('Manufacturing')
    ftlb.create_subtotals_format(workbook=workbook, worksheet=worksheetP, thedf=orderslist[0], timinglist=timingtest)
    ftlb.create_subtotals_format(workbook=workbook, worksheet=worksheetM, thedf=orderslist[1], timinglist=timingtest)
    worksheetT = workbook.add_worksheet('Timeline')
    ftlb.create_timeline_worksheet(workbook, worksheetT, demand)
    worksheetN = workbook.add_worksheet('PartsWithNoBOM')
    noboms = missingboms.get_parts()
    ftlb.create_miss_bom_worksheet(worksheetN, noboms)
    worksheetL = workbook.add_worksheet('PartsWithTooManyBOMs')
    lotsboms = manyboms.get_parts()
    ftlb.create_too_many_boms_worksheet(worksheetL, lotsboms)
    workbook.close()
    print('*The forecast is done!*')

"""Run forecast with an order taken away. This does not pull fresh data!
    This does the same thing as run_normal-forecast"""
def run_remove_order_forecast():
    missingboms = fs.No_BOMs()
    manyboms = fs.Many_BOMs()
    remove_order = remove_orders(missingboms, manyboms)
    timingtest = ftlb.find_timing_issues(remove_order[0], remove_order[1])
    demand = ftlb.find_demand_driver(remove_order[0])
    phantoms = ftlb.get_phantom_orders(demand)
    workbook = xlsxwriter.Workbook('Z:\Python projects\Chris Forecast\Forecast\RemoveOrderForecast.xlsx')
    orderslist = ftlb.split_phantoms(phantoms)
    worksheetP = workbook.add_worksheet('Purchasing')
    worksheetM = workbook.add_worksheet('Manufacturing')
    ftlb.create_subtotals_format(workbook=workbook, worksheet=worksheetP, thedf=orderslist[0], timinglist=timingtest)
    ftlb.create_subtotals_format(workbook=workbook, worksheet=worksheetM, thedf=orderslist[1], timinglist=timingtest)
    worksheetT = workbook.add_worksheet('Timeline')
    ftlb.create_timeline_worksheet(workbook, worksheetT, demand)
    worksheetN = workbook.add_worksheet('PartsWithNoBOM')
    noboms = missingboms.get_parts()
    ftlb.create_miss_bom_worksheet(worksheetN, noboms)
    worksheetL = workbook.add_worksheet('PartsWithTooManyBOMs')
    lotsboms = manyboms.get_parts()
    ftlb.create_too_many_boms_worksheet(worksheetL, lotsboms)
    workbook.close()
    print('*The forecast is done!*')

"""Run forecast with a tiered list of parts.  The tiered list helps prevent orders being attributed to the wrong "Grandparents".
   This pulls up to date info from Fishbowl, runs it, and saves it to an excel file."""
def run_normal_forecast_tiers_v2(ignore_schedule_errors=False):
    sql = ForecastAPI.run_queries() #Runs the function in ForecastAPI that pulls data from Fishbowl
    print(sql) #Prints Queries Successful!

    datalist = data_prep()
    """
    returns:
    datalist[0] = newordersdf
    datalist[1] = invdf
    datalist[2] = bomsdf
    """

    """ If schedule issues aren't your thing, pass ignore_schedule_errors=True when you call the funtion. """
    if ignore_schedule_errors==True:
        print('adjusting schedule to avoid issues ...')
        datalist[0] = bandaid_schedule_issues(datalist[0])  # <----- This is the only unique line to avoid trick the forecast.
                                                            #        It bumps orders that increase inventory about 5 years in the past so they resolve first.
                                                            #        It's not perfect, I think it might make excess schedule issues for make items.  Haven't proven it yet.

    normal_orders = [datalist[0], datalist[1]]  # newordersdf, invdf
    invdf = datalist[1]
    startinginvdf = datalist[1].copy() # invdf is going to change throughout, this is a reference for adding and inventory counter to the timeline later
    bomsdf = datalist[2]
    mypartsdf = gather_parts()

    print('going into create_bom_tiers')
    mytierlist = ftlb.create_bom_tiers_v2(bomsdf, mypartsdf)
    print('out of create_bom_tiers')


    missingboms = fs.No_BOMs()  # Creates an instance of the class in Forecast Settings No_BOMs
    manyboms = fs.Many_BOMs()  # Creates an instance of the class in Forecast Settings Many_BOMs
    tier = 1
    while len(mytierlist) > 0:
        print('Running tier', tier)
        tierlist = mytierlist[tier]
        normal_orders = complete_orders_loop_redux(tierlist, normal_orders[0], invdf, bomsdf, missingboms, manyboms)  # Runs the function that actually builds out the timeline
        normal_orders[0] = normal_orders[0].append(normal_orders[2])
        del mytierlist[tier]
        tier += 1


    print('*Timeline Has Been Created*')

    timingtest = ftlb.find_timing_issues(normal_orders[0], normal_orders[1])  # Finds parts that have phantom orders and are left with a positive inventory

    demand = ftlb.find_demand_driver(normal_orders[0])  # Runs the loop that figures out the top level driver for all the orders
    phantoms = ftlb.get_phantom_orders(demand)  # Returns all the phantom orders

    writer = pd.ExcelWriter('Z:\Python projects\Chris Forecast\Forecast\mytimelinetest.xlsx')  # Creates a test excel file
    demand.to_excel(writer, 'Sheet')  # Fills the test excel with the whole timeline
    writer.save()

    demand = ftlb.add_inv_counter(inputTimeline=demand, backdate='1999-12-31 00:00:00', invdf=startinginvdf)  # This adds the column that shows actual inventory of a part through its orders on the timeline

    workbook = xlsxwriter.Workbook('Z:\Python projects\Chris Forecast\Forecast\RegularForecast.xlsx')  # Create the actual final product excel file
    orderslist = ftlb.split_phantoms(phantoms)  # Seperates purchase and manufacture phantom orders
    worksheetP = workbook.add_worksheet('Purchasing')  # Create the seperate worksheets for purchasing and manufacturing
    worksheetM = workbook.add_worksheet('Manufacturing')

    """experimental code section, trying to add part descriptions to MFG worksheet"""
    descdf = pd.read_excel('Z:\Python projects\Chris Forecast\Forecast\ForecastRedoux\RawData\Descs.xlsx', header=0)
    orderslist[0] = pd.merge(orderslist[0].copy(), descdf.copy(), how='left', on='PART')
    orderslist[1] = pd.merge(orderslist[1].copy(), descdf.copy(), how='left', on='PART')
    columnList = ['ORDER', 'ITEM', 'ORDERTYPE', 'PART', 'QTYREMAINING', 'DATESCHEDULED', 'PARENT', 'Make/Buy', 'GRANDPARENT', 'DESCRIPTION']
    orderslist[0] = orderslist[0][columnList]
    orderslist[1] = orderslist[1][columnList]
    """seems to be working ... keeping it."""

    ftlb.create_subtotals_format(workbook=workbook, worksheet=worksheetP, thedf=orderslist[0], timinglist=timingtest)  # This function creates the subtotals in the worksheet
    ftlb.create_subtotals_format(workbook=workbook, worksheet=worksheetM, thedf=orderslist[1], timinglist=timingtest)  # Then is calls another function to add the data in FTLB
    worksheetT = workbook.add_worksheet('Timeline')  # Adds another worksheet for the timeline

    demand = pd.merge(demand.copy(), descdf.copy(), how='left', on='PART')  # Adding descriptions to timeline
    demand.sort_values(by=['PART', 'DATESCHEDULED'], ascending=[True, True], inplace=True)  #  Resorting it for readability before saving
    newColHeaders = ['ORDER', 'ITEM', 'ORDERTYPE', 'PART', 'QTYREMAINING', 'INV', 'DATESCHEDULED', 'PARENT', 'Make/Buy', 'GRANDPARENT', 'DESCRIPTION']
    demand = demand[newColHeaders]

    ###
    ftlb.create_timeline_worksheet(workbook, worksheetT, demand)  # Fills the timeline worksheet with data
    ###

    worksheetN = workbook.add_worksheet('PartsWithNoBOM')  # Adds another worksheet for No BOMs
    noboms = missingboms.get_parts()  # Returns the parts with no BOMs
    ftlb.create_miss_bom_worksheet(worksheetN, noboms)  # Fills the worksheet with data
    worksheetL = workbook.add_worksheet('PartsWithTooManyBOMs')  # Creates another worksheet for too many BOMs
    lotsboms = manyboms.get_parts()  # Returns the parts with many active BOMs
    ftlb.create_too_many_boms_worksheet(worksheetL, lotsboms)  # Fills the worksheet with data
    workbook.close()  # Closing the workbook commits all the data

    # Saving a copy of the end timeline as "demand.xlsx" to use in other projects.
    writer = pd.ExcelWriter('Z:\Python projects\Chris Forecast\Forecast\demand.xlsx')
    demand.to_excel(writer, 'timeline')
    writer.save()

    print('*The forecast is done!*')


