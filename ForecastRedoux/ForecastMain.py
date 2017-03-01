__author__ = 'Chris'

import ForecastAPI
import os
import pandas as pd
import ForecastTimelineBackend as ftlb
import ForecastSettings as fs
import xlsxwriter
import datetime as dt
import ForecastEmail as fe

"""Pull the orders data"""
def gather_orders():
    mopath = os.path.join('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'MOs.xlsx')
    modf = pd.read_excel(mopath, header=0)
    orgdf = ftlb.create_mo_parents(modf)
    popath = os.path.join('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'POs.xlsx')
    tempdf = pd.read_excel(popath, header=0)
    orgdf = orgdf.append(tempdf.copy())
    sopath = os.path.join('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'SOs.xlsx')
    tempdf = pd.read_excel(sopath, header=0)
    orgdf = orgdf.append(tempdf.copy())
    orgdf = orgdf.sort_values(by=['PART', 'DATESCHEDULED'], ascending=[True, False])
    return orgdf

"""Pull the parts data"""
def gather_parts():
    partspath = os.path.join('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'Parts.xlsx')
    partsdf = pd.read_excel(partspath)
    partsdf = partsdf.sort_values(by='PART', ascending=True)
    return partsdf

"""Pull the inventory data"""
def gather_inventory():
    invpath = os.path.join('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'INVs.xlsx')
    orgdf = pd.read_excel(invpath, header=0)
    orgdf = orgdf.sort_values(by='PART', ascending=True)
    return orgdf

"""Pull the BOM data"""
def gather_boms():
    bompath = os.path.join('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/ForecastRedoux/RawData', 'BOMs.xlsx')
    orgdf = pd.read_excel(bompath, header=0)
    orgdf = orgdf.sort_values(by=['BOM','FG','PART'], ascending=[True,True,True])
    return orgdf

"""Create make/buy table."""
def make_buy(mbdf):
    mbdf = mbdf.drop('AvgCost', axis=1)
    mbdf = mbdf.sort_values(by='PART', ascending=True)
    mbdf = mbdf.drop_duplicates('PART')
    return mbdf

"""Add make/buy column to main table."""
def add_make_buy(orgdf, mbdf):
    newdf = pd.merge(orgdf, mbdf, on='PART', how='left')
    return newdf

"""Adds new orders to original orders"""
def add_new_orders(ordersdf, newordersdf):
    allordersdf = ordersdf.append(newordersdf)
    allordersdf.reset_index(drop=True, inplace=True)
    return allordersdf

"""Just combining all the data pulling functions."""
def data_prep():
    ordersdf = gather_orders()
    partsdf = gather_parts()
    mbdf = make_buy(partsdf)
    newordersdf = add_make_buy(ordersdf, mbdf)
    invdf = gather_inventory()
    bomsdf = gather_boms()
    return [newordersdf, invdf, bomsdf]

"""Loops to capture all orders"""
def complete_orders_loop_redoux(missingboms, manyboms):
    ordersdf = gather_orders()
    partsdf = gather_parts()
    mbdf = make_buy(partsdf)
    newordersdf = add_make_buy(ordersdf, mbdf)
    invdf = gather_inventory()
    bomsdf = gather_boms()
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

"""Run forecast on fake order."""
def run_fake_orders(missingboms, manyboms):
    newbuilds = pd.read_excel('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/AdditionalInfo/PartsToBuild.xlsx', sheetname='Sheet1')
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

"""Takes away any orders that are known to not be included in the demand for a part."""
def remove_orders(missingboms, manyboms):
    droporders = pd.read_excel('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/AdditionalInfo/OrdersToRemove.xlsx', sheetname='Sheet1')
    ordersdf = gather_orders()
    partsdf = gather_parts()
    mbdf = make_buy(partsdf)
    newordersdf = add_make_buy(ordersdf, mbdf)
    invdf = gather_inventory()
    bomsdf = gather_boms()
    for item, row in droporders.iterrows():
        if row['OrderType'] == 'MO':
            newtimeline = newordersdf.ix[newordersdf['ITEM'] != row['OrderNum']].copy()
        else:
            newtimeline = newordersdf.ix[newordersdf['ORDER'] != row['OrderNum']].copy()
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

"""Run the normal forecast"""
def run_normal_forecast():
    sql = ForecastAPI.run_queries()
    print(sql)
    missingboms = fs.No_BOMs()
    manyboms = fs.Many_BOMs()
    normal_orders = complete_orders_loop_redoux(missingboms, manyboms)
    print('*Timeline Has Been Created*')
    timingtest = ftlb.find_timing_issues(normal_orders[0], normal_orders[1])
    demand = ftlb.find_demand_driver(normal_orders[0])
    phantoms = ftlb.get_phantom_orders(demand)
    writer = pd.ExcelWriter('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/mytimelinetest.xlsx')
    demand.to_excel(writer, 'Sheet')
    writer.save()
    workbook = xlsxwriter.Workbook('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/RegularForecast.xlsx')
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
    email_forecast()
    print('*The forecast is done!*')

"""Run forecast with an added order. This does not pull fresh data!"""
def run_add_order_forecast():
    missingboms = fs.No_BOMs()
    manyboms = fs.Many_BOMs()
    add_order = run_fake_orders(missingboms, manyboms)
    timingtest = ftlb.find_timing_issues(add_order[0], add_order[1])
    demand = ftlb.find_demand_driver(add_order[0])
    phantoms = ftlb.get_phantom_orders(demand)
    workbook = xlsxwriter.Workbook('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/AddOrderForecast.xlsx')
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

"""Run forecast with an order taken away. This does not pull fresh data!"""
def run_remove_order_forecast():
    missingboms = fs.No_BOMs()
    manyboms = fs.Many_BOMs()
    remove_order = remove_orders(missingboms, manyboms)
    timingtest = ftlb.find_timing_issues(remove_order[0], remove_order[1])
    demand = ftlb.find_demand_driver(remove_order[0])
    phantoms = ftlb.get_phantom_orders(demand)
    workbook = xlsxwriter.Workbook('/mnt/manufacturing/Shared Services/Projects/RaspberryPi/Forecast/RemoveOrderForecast.xlsx')
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

"""Create and send the email."""
def email_forecast():
    today = dt.datetime.now().strftime("%Y-%d-%B %I%M%p")
    fe.send_email(today)