__author__ = 'Chris'

import pandas as pd
import ForecastBOMExploder as bom
import datetime as dt
import ForecastSettings as fs

"""Creating proper parents for MOs"""
def create_mo_parents(orgdf):
    molist = orgdf.MOID.unique()
    for each in molist:
        tempdf = orgdf.ix[orgdf['MOID'] == each].copy()
        tempfindf = tempdf.ix[tempdf['ORDERTYPE'] == 'Finished Good'].copy()
        temprawdf = tempdf.ix[tempdf['ORDERTYPE'] == 'Raw Good'].copy()
        for index, row in tempdf.iterrows():
            tempfinpart = tempfindf.ix[tempfindf['ORDER'] == row['ORDER']].copy()
            if not tempfinpart.empty:
                tempfinpart.reset_index(drop=True, inplace=True)
                temprawpart = temprawdf.ix[temprawdf['PART'] == tempfinpart.ix[0, 'PART']].copy()
                if not temprawpart.empty:
                    tempparent = temprawdf.ix[temprawdf['PART'] == tempfinpart.ix[0, 'PART']]
                    tempparent.reset_index(drop=True, inplace=True)
                    tempparent = tempparent.loc[0, 'ORDER']
                    testparent = tempfindf.ix[tempfindf['ORDER'] == tempparent].copy()
                    newparent = testparent.reset_index(drop=True, inplace=False)
                    testrawparent = temprawdf.ix[temprawdf['PART'] == newparent.loc[0, 'PART']].copy()
                    if testrawparent.empty:
                        orgdf.loc[index, 'PARENT'] = tempparent
                    else:
                        parentindex = testparent.index.tolist()
                        tempnewparent = orgdf.loc[parentindex[0], 'PARENT']
                        orgdf.loc[index, 'PARENT'] = tempnewparent
    orgdf = orgdf.drop('MOID', 1)
    return orgdf

"""Takes orders and inv and finds the first shortages"""
def find_next_shortage_redoux(invdf, postordersdf):
    postordersdf.reset_index(drop=True, inplace=True)
    partslist = postordersdf.PART.unique()
    shortagedf = pd.DataFrame()
    column_headers = ['ORDER', 'ITEM', 'ORDERTYPE', 'PART', 'QTYREMAINING', 'DATESCHEDULED', 'PARENT', 'Make/Buy']
    preordersdf = pd.DataFrame(columns=column_headers)
    for each in partslist:
        tempordersdf = postordersdf.ix[postordersdf['PART'] == each].copy()
        tempordersdf['DATESCHEDULED'] = pd.to_datetime(tempordersdf['DATESCHEDULED'])
        tempordersdf = tempordersdf.sort_values(by=['DATESCHEDULED', 'QTYREMAINING', 'ORDER'], ascending=[True, False, True])
        currentinv = invdf.loc[invdf['PART'] == each]
        if currentinv.empty:
            workinginv = 0
            tempinvdf = pd.DataFrame({'PART': [each], 'INV': [0]})
            invdf = invdf.append(tempinvdf)
        else:
            currentinv.reset_index(drop=True, inplace=True)
            workinginv = currentinv.at[0,'INV']
        for index, row in tempordersdf.iterrows():
            if (workinginv <= 0 and row['QTYREMAINING'] < 0):
                shortagedf = shortagedf.append(row)
                workinginv = workinginv + row['QTYREMAINING']
                preordersdf = preordersdf.append(tempordersdf.ix[index])
                postordersdf = postordersdf.drop(labels=index)
                break
            else:
                workinginv = workinginv + row['QTYREMAINING']
                preordersdf = preordersdf.append(tempordersdf.ix[index])
                postordersdf = postordersdf.drop(labels=index)
                if (workinginv < 0 and row['QTYREMAINING'] < 0):
                    row['QTYREMAINING'] = workinginv
                    shortagedf = shortagedf.append(row)
                    break
        invdf.ix[(invdf['PART'] == each),'INV'] = workinginv
    return [shortagedf, postordersdf, preordersdf, invdf]

"""Makes new orders to put back into timeline."""
def make_new_orders(shortordersdf, bomsdf, missingbomlist, manybomlist):
    column_headers = ['ORDER', 'ITEM', 'ORDERTYPE', 'PART', 'QTYREMAINING', 'DATESCHEDULED', 'PARENT', 'Make/Buy']
    newordersdf = pd.DataFrame(columns=column_headers)
    for index, row in shortordersdf.iterrows():
        if row['ITEM'] == 'Imaginary':
            order = fs.Index().show_fake_index()
        else:
            order = fs.Index().show_index()
        if row['ITEM'] == 'Imaginary':
            item = 'Imaginary'
        else:
            item = 'Phantom'
        if row['Make/Buy'] == 'Make':
            ordertype = 'Finished Good'
            bomdf = bom.order_exploder(bomsdf, row, order, missingbomlist, manybomlist)
            newordersdf = newordersdf.append(bomdf)
        else:
            ordertype = 'Purchase'
        part = row['PART']
        qty = (row['QTYREMAINING'] * -1)
        date = pd.to_datetime(row['DATESCHEDULED']) - dt.timedelta(days=1)
        date = date.date()
        parent = row['ORDER']
        mb = row['Make/Buy']
        tempdatalist = [order, item, ordertype, part, qty, date, parent, mb]
        tempdic = {}
        for i in range(0, 8):
            tempdic[column_headers[i]] = [tempdatalist[i]]
        tempdf = pd.DataFrame.from_dict(tempdic)
        newordersdf = newordersdf.append(tempdf)
    return newordersdf

"""Create what if order."""
def create_phantom_order(bompart, qty, completiondate, bomsdf, missingbomlist, manybomlist):
    column_headers = ['ORDER', 'ITEM', 'ORDERTYPE', 'PART', 'QTYREMAINING', 'DATESCHEDULED', 'PARENT', 'Make/Buy']
    neworderdf = pd.DataFrame(columns=column_headers)
    order = fs.Index().show_fake_index()
    item = 'Imaginary'
    ordertype = 'Finished Good'
    parent = order
    mb = 'Make'
    tempdatalist = [order, item, ordertype, bompart, qty, completiondate, parent, mb]
    tempdic = {}
    for i in range(0, 8):
        tempdic[column_headers[i]] = [tempdatalist[i]]
    bompartdf = pd.DataFrame.from_dict(tempdic)
    neworderdf = neworderdf.append(bompartdf)
    bompartdf['DATESCHEDULED'] = bompartdf['DATESCHEDULED'] + dt.timedelta(days=1)
    bomdf = bom.order_exploder_new_order(bomsdf, bompartdf, missingbomlist, manybomlist)
    neworderdf = neworderdf.append(bomdf)
    return neworderdf

"""Add what if order to timeline"""
def add_phantom_order(phantomdf, timeline):
    workingtimeline = timeline.append(phantomdf)
    return workingtimeline

"""Checks ending inventory to see about timing issues."""
def find_timing_issues(timeline, lastinv):
    phantomtimeline = timeline.ix[timeline['ITEM'] == 'Phantom']
    phantomtimelineP = phantomtimeline.ix[phantomtimeline['ORDERTYPE'] == 'Purchase'].copy()
    phantomtimelineF = phantomtimeline.ix[phantomtimeline['ORDERTYPE'] == 'Finished Good'].copy()
    phantomtimeline = phantomtimelineF.append(phantomtimelineP)
    partlist = phantomtimeline.PART.unique()
    phantominv = lastinv[lastinv.PART.isin(partlist)]
    timinglist = phantominv.ix[phantominv['INV'] > 0]
    timinglist = timinglist['PART'].tolist()
    return timinglist

"""This loops through finding the original demand driver."""
def find_demand_driver(timeline):
    phantomtimeline = timeline.ix[timeline['ITEM'] == 'Phantom']
    for index, row in phantomtimeline.iterrows():
        if row['ORDERTYPE'] == 'Purchase':
            tempparent = timeline.ix[timeline['ORDER'] == row['PARENT']].copy()
            tempparent = tempparent.ix[tempparent['PART'] == row['PART']].copy()
            while True:
                tempparent.reset_index(drop=True, inplace=True)
                if tempparent.get_value(0, 'ITEM') == 'Phantom':
                    if tempparent.get_value(0, 'ORDERTYPE') == 'Raw Good':
                        tempparenttemp = timeline.ix[timeline['ORDER'] == tempparent.get_value(0, 'ORDER')]
                        tempparent = tempparenttemp.ix[tempparenttemp['ORDERTYPE'] == 'Finished Good']
                    else:
                        tempparenttemp = timeline.ix[timeline['ORDER'] == tempparent.get_value(0, 'PARENT')]
                        tempparent = tempparenttemp.ix[tempparenttemp['PART'] == tempparent.get_value(0, 'PART')]
                else:
                    tempparent.reset_index(drop=True, inplace=True)
                    rowp = tempparent.get_value(0, 'PARENT')
                    timeline.ix[index, 'GRANDPARENT'] = rowp
                    break
        else:
            if row['ORDERTYPE'] == 'Raw Good':
                tempparenttemp = timeline.ix[timeline['ORDER'] == row['ORDER']]
                tempparent = tempparenttemp.ix[tempparenttemp['ORDERTYPE'] == 'Finished Good']
            else:
                tempparent = timeline[timeline['ORDER'] == row['ORDER']]
                tempparent = tempparent[tempparent['PART'] == row['PART']]
            while True:
                tempparent.reset_index(drop=True, inplace=True)
                if tempparent.get_value(0, 'ORDERTYPE') == 'Raw Good':
                    tempparenttemp = timeline.ix[timeline['ORDER'] == tempparent.get_value(0, 'ORDER')]
                    tempparent = tempparenttemp.ix[tempparenttemp['ORDERTYPE'] == 'Finished Good']
                    tempparent.reset_index(drop=True, inplace=True)
                if tempparent.get_value(0, 'ITEM') == 'Phantom':
                    tempparenttemp = timeline.ix[timeline['ORDER'] == tempparent.get_value(0, 'PARENT')]
                    tempparent = tempparenttemp.ix[tempparenttemp['PART'] == tempparent.get_value(0, 'PART')]
                else:
                    timeline.ix[index, 'GRANDPARENT'] = tempparent.get_value(0, 'PARENT')
                    break
    return timeline

"""Gather Phantom parts."""
def get_phantom_orders(thedf):
    phantomdf = thedf.ix[thedf['ITEM'] == 'Phantom']
    phantomdf = phantomdf.sort_values(by=['PART', 'DATESCHEDULED', 'ORDER'], ascending=[True, True, True])
    return phantomdf

"""Fills the cells in the excel workbook."""
def fill_subtotals_workbook(worksheet, thedf, format):
    columns = list(thedf.columns.values)
    for each in columns:
        worksheet.write(0, columns.index(each), each)
    rowcounta = 1
    uniqueA = thedf.PART.unique()
    for item in uniqueA:
        rowcountb = 0
        tempAdf = thedf.ix[thedf['PART'] == item]
        tempAdf.reset_index(drop=True, inplace=True)
        while rowcountb < len(tempAdf):
            for each in columns:
                value = tempAdf.get_value(rowcountb, each)
                if each == 'DATESCHEDULED':
                    worksheet.write_datetime((rowcounta + rowcountb), columns.index(each), value, format)
                else:
                    worksheet.write((rowcounta + rowcountb), columns.index(each), value)
            rowcountb += 1
        rowcountat = rowcounta + rowcountb
        worksheet.write(rowcountat, 3, item)
        rowcounta += 1
        worksheet.write_formula(rowcountat, 4, '=SUBTOTAL(9, E%s:E%s)' %(rowcounta, rowcountat))
        rowcounta = rowcountat + 1

"""Creates the framework in the excel workbook."""
def create_subtotals_format(workbook, worksheet, thedf, timinglist):
    rowcounta = 1
    rowcountb = 1
    uniqueA = thedf.PART.unique()
    for item in uniqueA:
        tempAdf = thedf.ix[thedf['PART'] == item]
        tempAdf.reset_index(drop=True, inplace=True)
        for line in range(0, len(tempAdf)):
            newline = line + rowcountb
            worksheet.set_row(newline, None, None, {'hidden': True, 'level': 2})
            rowcounta += 1
        if item in timinglist:
            timingformat = workbook.add_format()
            timingformat.set_bg_color('pink')
            worksheet.set_row((rowcounta), None, timingformat, {'level': 1, 'collapsed': True})
        else:
            worksheet.set_row((rowcounta), None, None, {'level': 1, 'collapsed': True})
        rowcounta += 1
        rowcountb = rowcounta
    dateformat = workbook.add_format({'num_format': 'mm/dd/yy'})
    fill_subtotals_workbook(worksheet, thedf, dateformat)

"""Create Timeline Worksheet."""
def create_timeline_worksheet(workbook, worksheet, timeline):
    dateformat = workbook.add_format({'num_format': 'mm/dd/yy'})
    timeline.reset_index(drop=True, inplace=False)
    timeline = timeline.fillna(0)
    columns = list(timeline.columns.values)
    columnindex = len(columns)
    for each in range(0, columnindex):
        worksheet.write(0, each, columns[each])
    for row in range(1, len(timeline)):
        for column in range(0, columnindex):
            if column == 5:
                value = timeline.get_value(row, column, True)
                worksheet.write_datetime(row, column, value, dateformat)
            else:
                value = timeline.get_value(row, column, True)
                worksheet.write(row, column, value)

"""Splits up all the phantom orders for purchasing and manufacturing."""
def split_phantoms(phantomdf):
    purchasedf = phantomdf.ix[phantomdf['ORDERTYPE'] == 'Purchase'].copy()
    manufacturingdf = phantomdf.ix[phantomdf['ORDERTYPE'] == 'Finished Good'].copy()
    return [purchasedf, manufacturingdf]

"""Create dataframe from missing boms"""
def miss_bom_df(misslist):
    temp = {'Parts with no BOM' : misslist}
    missdf = pd.DataFrame(temp)
    return missdf

"""Creates sheet for missing boms"""
def create_miss_bom_worksheet(worksheet, parts):
    worksheet.write(0, 0, 'PartsWithNoBOM')
    row = 1
    for each in parts:
        worksheet.write(row, 0, each)
        row += 1

"""Create dataframe from too many boms"""
def many_bom_df(manylist):
    temp = {'Parts with too many BOMs' : manylist}
    manydf = pd.DataFrame(temp)
    return manydf

"""Creates sheet for too many boms"""
def create_too_many_boms_worksheet(worksheet, parts):
    worksheet.write(0, 0, 'PartsWithTooManyBOMs')
    row = 1
    for each in parts:
        worksheet.write(row, 0, each)
        row += 1