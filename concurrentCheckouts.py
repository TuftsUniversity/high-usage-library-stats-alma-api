############################################################################################
############################################################################################
########
########    Title:      concurrentCheckouts
########    Author:     Henry Steele, Library Technology Services, Tufts University
########    Date:       June 2018
########
########    This work is licensed under the Creative Commons Attribution-NonCommercial 4.0 International License.
########    To view a copy of this license, visit http://creativecommons.org/licenses/by-nc/4.0/
########    or send a letter to Creative Commons, PO Box 1866, Mountain View, CA 94042, USA.
########
########    Purpose:
########        Create a report of concurrent checkouts that occured on multiple
########        copies of the same volume, based on an exporte Analytics report with
########        the criteria below.   Note the required format.
########
########        This script finds out how often during the given time periods
########        that multiple copies of the same volume were out at the same time,
########        and how often that all copies of the same volume were out at the same time
########
########        This report assumes the Tufts University rubric for multiple copies,
########        that they will have the same MMS Id and call number, but different barcodes
########
########        The report returns counts for when all copies of a title were out
########        at the same time, but excludes these counts if there is only one copy of a title
########
########    Input:
########        The Analtyics report should have the following fields.  They can be in any
########        order, and you can have additional fields (they'll be ignored) but the field names
########        must be as below.  It should be in Excel format .xlsx format
########
########        - fulfilllment table with at least
########            + Title
########            + MMS Id
########            + Permanent Call Number
########            + Barcode
########            + Loan Date
########            + Loan Time
########            + Return Date
########            + Return Time
########
########    Dependencies.  Note that this code is currently configured for Python 2.7, but I've noted in
########    the dependencies below and in various places in the code how to convert (refactor) this for Python > 3
########
########        - Python 2.7
########        - use pip or another Python installation utility to install:
########            + pandas (this also installs numpy)
########            + tkFileDialog
########            + xlwt
########            + xlsxwriter
########            + xlrd
########                + you'll also need to intall xlrd for read_exce in pandas to work
########
########       - Python > 3.0
########            + pandas (this also installs numpy)
########            + tkinter
########            + xlwt
########            + xlsxwriter
########
########    Output:
########        The script will output an Excel workbook of concurrent checkouts counts
########        for each volume.
########
########    Method:
########
########        Dataframe "a" is a parsed version of the input report from Analytics.
########
########        It contains 'Title', 'MMS Id', 'Permanent Call Number, 'Barcode',
########        'Loan Datetime', 'Return Datetime'
########        This is used to compare loan periods for different items of the same volume
########
########        The logic of the script is to load loan and return times into
########        dataframe "c", where each datetime in which either a loan or return occurred
########        is a column in the dataframe, and each copy (barcode) of the same volume
########        is a row.  In the cell for each row,column, the script records whether it
########        was a loan or a return.
########
########        Dataframe c is rearranged by column name, so that the column names (datetimes)
########        for loans and returns are in order.
########
########        With the columns arranged in this way, some loans will span multiple columns,
########        i.e., another transacation's loan or return will have occured in the middle
########        of the loan period of this transacation. This is the kind of event the script
########        is looking for, because it means loan periods of different copies of the same
########        volume overlapped.  To be able to analyze this situation, the script fills
########        in the columns between "loan" and "return" for transacations that span multiple
########        columns with "on loan."
########
########        The last step the script takes is to analyze each column (datetime)
########        to see how many of the volume's copies were on loan at that time.
########        The important question the output will answer how often all copies of a given volumes
########        were out at a given time because that may indicate that the library
########        needs to purchase more copies, depending on how often this happened

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import time

# for >= Python 2.7 < Python 3
from tkFileDialog import askopenfilename

#for Python 3
# from tkinter import Tk
# from tkinter import filedialog
# from tkinter import *


import os
import xlwt
import xlsxwriter
import openpyxl
import re

# Python 2
from django.utils.encoding import smart_str, smart_unicode

# Python 3
# from django.utils.encoding import smart_str

def utf8(title):
    return title.decode('utf-8')


# for >= Python 2.7 < Python 3
filename = askopenfilename(title = "Select EXCEL file containing TRANSACTIONS")
# for Python 3
#filenameCirc = filedialog.askopenfilename(title = "Select EXCEL file containing TRANSACTIONS")



print(filename + "\n")


# create output directory
oDir = "./Output"
if not os.path.isdir(oDir) or not os.path.exists(oDir):
       os.makedirs(oDir)

# create dataframe
cc = pd.read_excel(filename, header=2, encoding = 'windows-1252', dtype={'MMS Id': 'str', 'Permanent Call Number': 'str', 'Barcode': 'str'}, converters={'Loan Date': pd.to_datetime, 'Loan Time': pd.to_datetime, 'Return Date': pd.to_datetime, 'Return Time': pd.to_datetime}, skipfooter=1)

todays_date = datetime.now()
current_time = datetime.now()

cc['Return Date'] = cc['Return Date'].fillna(todays_date)

cc['Return Time'] = cc['Return Time'].fillna(current_time)


cc['Loan Date'] = cc['Loan Date'].apply(lambda x: x.strftime('%m-%d-%Y'))
print(cc)
cc['Loan Time'] = cc['Loan Time'].apply(lambda x: x.strftime('%H:%M:%S'))
cc['Return Date'] = cc['Return Date'].apply(lambda x: x.strftime('%m-%d-%Y'))
cc['Return Time'] = cc['Return Time'].apply(lambda x: x.strftime('%H:%M:%S'))


print(cc)
cc['Loan Datetime'] = pd.to_datetime(cc['Loan Date'] + ' ' + cc['Loan Time'])
cc['Return Datetime'] = pd.to_datetime(cc['Return Date'] + ' ' + cc['Return Time'])


# sort so that the script can loop through volumes sequentially

cc = cc.sort_values(['MMS Id','Permanent Call Number', 'Barcode', 'Loan Datetime', 'Return Datetime'])





dd  = pd.DataFrame()

ee = pd.DataFrame()

output_excel_file = pd.ExcelWriter(oDir + '/Output Dataframes.xlsx', engine='xlsxwriter')
writerAll = pd.ExcelWriter(oDir + '/Counts.xlsx', engine='xlsxwriter')

# counter for entire sheet
x = 0

# counter for number of volumes
volumeCount = 0





workbook = writerAll.book


totalBarcodeCount = 0
# loop through master dataframe

while x < len(cc):
    volumeCount += 1
    # y and count are the counter for looping through transactions for each volume
    y = x
    count = 0

    title = cc.iloc[x]['Title']
    mms_id = cc.iloc[x]['MMS Id']
    call_number = cc.iloc[x]['Permanent Call Number']


    columns = ['Title', 'MMS Id', 'Permanent Call Number', 'Barcode', 'Loan Datetime', 'Loan Date', 'Return Datetime', 'Return Date']
    a = pd.DataFrame(columns=columns)


    d = {}


    # populate the dataframes and the series for each title
    a = a.append({'Title':cc.iloc[x]['Title'], 'MMS Id': cc.iloc[x]['MMS Id'], 'Call Number': cc.iloc[x]['Permanent Call Number'], 'Barcode': cc.iloc[x]['Barcode'], 'Loan Datetime': cc.iloc[x]['Loan Datetime'], 'Loan Date': cc.iloc[x]['Loan Date'], 'Return Datetime': cc.iloc[x]['Return Datetime'], 'Return Date': cc.iloc[x]['Return Date']}, ignore_index=True)


    y += 1
    count += 1

    while y < len(cc) and cc.iloc[y]['MMS Id'] == cc.iloc[y - 1]['MMS Id'] and cc.iloc[y]['Permanent Call Number'] == cc.iloc[y - 1]['Permanent Call Number']:

        a = a.append({'Title':cc.iloc[y]['Title'], 'MMS Id': cc.iloc[y]['MMS Id'], 'Call Number': cc.iloc[y]['Permanent Call Number'], 'Barcode': cc.iloc[y]['Barcode'], 'Loan Datetime': cc.iloc[y]['Loan Datetime'], 'Loan Date': cc.iloc[y]['Loan Date'], 'Return Datetime': cc.iloc[y]['Return Datetime'], 'Return Date': cc.iloc[y]['Return Date']}, ignore_index=True)

        y = y + 1
        count = count + 1


    # count is the number of transacations on this volume
    # z is the counter for the transacations on this volume
    # concurrent count records how many times more than one copy on this volume was checked out
    # f is the counter for comparing transacations within the same barcode
    # maxedOutCount records the number of times all copies of the volume were checked out, excluding single copies
    # barcode dict is a list of the barcodes on a given volume, that will be used to rename the rows of the temporary dataframe

    z = 0



    barcodeDict = {}

    barcodeCount = 0
    transactionWithinBarcodeCount = 0
    c = pd.DataFrame()
	# dataframe c plots each loan or return as a location by barcode and datetime
    # the row is barcode, and there is a separate column for each datetime, whether it's loan or return

    # the columns are resorted at the end to make sure they occur in sequential order by datetime, even though
    # they will be close to in order because of the initial sorting on the master dataframe

    while z < count:

        f = z + 1

        firstLoanIndex = str(a.at[z, 'Loan Datetime'])
        barcode = str(a.iloc[z]['Barcode'])

        if firstLoanIndex in c:
            firstLoanIndex += ":0" + str(z)

        c.insert(loc=transactionWithinBarcodeCount, column=firstLoanIndex, value="")

        c.at[barcodeCount, firstLoanIndex] = "loan"
        transactionWithinBarcodeCount += 1





        firstReturnIndex = str(a.iloc[z]['Return Datetime'])

        if firstReturnIndex in c:
            firstReturnIndex += ":0" + str(z)

        c.insert(loc=transactionWithinBarcodeCount, column=firstReturnIndex, value="")
        c.at[barcodeCount, firstReturnIndex] = "return"
        transactionWithinBarcodeCount =+ 1

        while f < count and a.iloc[z]["Barcode"] == a.iloc[f]["Barcode"]:

            loanIndex = str(a.iloc[f]['Loan Datetime'])

            # handling collisions if two datetimes are at exactly the same date time (hours, minutes, and seconds)
            if loanIndex in c:
                loanIndex += ":01"

            c.insert(loc=transactionWithinBarcodeCount, column=loanIndex, value="")

            c.at[barcodeCount, loanIndex] = "loan"
            transactionWithinBarcodeCount += 1

            returnIndex = str(a.iloc[f]['Return Datetime'])
            if returnIndex in c:
                returnIndex += ":01"

            c.insert(loc=transactionWithinBarcodeCount, column=returnIndex, value="")


            c.at[barcodeCount, returnIndex] = "return"
            transactionWithinBarcodeCount += 1

            f += 1

        z += f
        barcodeDict[barcodeCount] = barcode

        c = c.rename(index=barcodeDict)
        barcodeCount += 1
        totalBarcodeCount += 1
    #c = c.fillna("blank")


    # for Python 3
    #c = c.reindex(sorted(c.columns), axis=1)
    # for >= Python 2.7 < Python 3
    c = c.reindex_axis(sorted(c.columns), axis=1)

    columnCount = len(c.columns)

    # fill NaN columns in the dataframe with the value 0. This makes it easier to search for


    # the loops below adds "on loan" to transacations that spanned multiple
    # datetime columns and adds the value "on loan" to cells in the middle
    # of transactions that span multiple columns
    l = 0
    m = 0
    if barcodeCount > 1:
        while l < barcodeCount:
            while m < columnCount:
                if c.iloc[l][m] == "loan":
                    d = 1
                    #date = c.at[l, m]
                    #print("Date: " + str(date) + "\n")
                    while m + d < columnCount and c.iloc[l][m + d] != "loan" and c.iloc[l][m + d] != "return":
                        c.iloc[l][m + d] = "on loan"
                        d += 1
                m += 1
            l += 1




    concurrentDates = {}

    maxedOutDates = {}

    # check for the number of times concurrent checkouts occurred
    concurrentCount = 0
    maxedOutCount = 0
    concurrentLoanRunCounter = 0
    maxedOutLoanRunCounter = 0
    concurrentDates['Title'] = title
    concurrentDates['MMS Id'] = mms_id
    concurrentDates['Call Number'] = call_number
    for column in c.columns:
        if len(c[c[column] == "loan"]) + len(c[c[column] == "on loan"])  > 1 and barcodeCount > 1:
            concurrentLoanRunCounter += 1

        elif concurrentLoanRunCounter > 0 and len(c[c[column] == "loan"]) + len(c[c[column] == "on loan"]) <= 1 and barcodeCount > 1:
            concurrentLoanRunCounter = 0
            concurrentCount += 1
            column = str(column) + '.' + str(volumeCount)
            concurrentDates[column] = 1


	#check for times all copies were checked out
    maxedOutDates['Title'] = title
    maxedOutDates['MMS Id'] = mms_id
    maxedOutDates['Call Number'] = call_number
    for column in c.columns:
        if len(c[c[column] == "loan"]) + len(c[c[column] == "on loan"]) == barcodeCount and barcodeCount > 1:
            maxedOutLoanRunCounter += 1

        elif maxedOutLoanRunCounter > 0 and len(c[c[column] == "loan"]) + len(c[c[column] == "on loan"]) < barcodeCount and barcodeCount > 1:
            maxedOutLoanRunCounter = 0
            maxedOutCount += 1
            column = str(column) + '.' + str(volumeCount)
            maxedOutDates[column] = 1





    if concurrentCount > 0:

        print("Concurrent date dict: \n")
        print(concurrentDates)
        dd = dd.append(concurrentDates, ignore_index=True)
        print("Concurrent with new columns\n")
        print(dd)

    if maxedOutCount > 0:

        print("All copies in use date dict: \n")
        print(maxedOutDates)
        ee = ee.append(maxedOutDates, ignore_index=True)
        print("Maxed out with new columns\n")
        print(ee)

    print("\n\n")
    print(c)
    print("MMS Id: " + str(mms_id) + " with barcodes: " + str(barcodeDict) + "\n")

    print("Concurrent checkouts count:                                         " + str(concurrentCount) + "\n")
    print("Concurrent checkout times:                                          " + str(concurrentDates) + "\n")
    print("All copies in use count:                                            " + str(maxedOutCount) + "\n")
    print("All copies in use times:                                            " + str(maxedOutDates) + "\n")
    print("\n\n")
    c.insert(loc=0, column='Title', value=title)
    c.insert(loc=1, column='Call Number', value=call_number)
    c.insert(loc=2, column='MMS Id', value=mms_id)
    c.insert(loc=3, column='Copy Count', value=barcodeCount)
    c.insert(loc=4, column='Loan Count', value=count)
    c.insert(loc=5, column="Concurrent Checkout Count", value = concurrentCount)
    c.insert(loc=6, column="All Copies in Use Count", value = maxedOutCount)

    # sheet_name_title = re.sub("[!@#$%^&*()[]{};:,./<>?\|`~-=_+]", " ", title)
    # c.to_excel(output_excel_file, sheet_name=str(mms_id), startrow=0, startcol=0, index=False)
    #
    # a.to_excel(output_excel_file, sheet_name="df A - " + str(mms_id), startrow=0, startcol=0, index=False)


    o = c.loc[:, ['Title', 'Call Number', 'MMS Id', 'Copy Count', 'Loan Count', 'Concurrent Checkout Count', 'All Copies in Use Count']]
    o = o.drop_duplicates()


    #startRow = totalBarcodeCount-barcodeCount
    print("Volume count: " + str(volumeCount) + "\n")
    if volumeCount - 1 == 0:
        o.to_excel(writerAll, sheet_name='Counts', startrow=volumeCount - 1, startcol=0, index=False)
    else:
        o.to_excel(writerAll, sheet_name='Counts', startrow=volumeCount, startcol=0, header=False, index=False)

    x += count










worksheet = writerAll.sheets['Counts']
# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 30)
worksheet.set_column('B:B', 30)
worksheet.set_column('C:C', 30)
worksheet.set_column('D:D', 15)
worksheet.set_column('E:E', 15)
worksheet.set_column('F:F', 30)
worksheet.set_column('G:G', 30)


# Green fill with dark green text.
green_format = workbook.add_format({'bg_color':   '#C6EFCE', 'font_color': '#006100'})
deep_green_format = workbook.add_format({'bg_color':   '#73c48b', 'font_color': '#006100'})



worksheet.conditional_format(1, 5, volumeCount + 1, 5, {'type': 'cell', 'criteria': '>', 'value': 0,'format': green_format})
worksheet.conditional_format(1, 6, volumeCount + 1, 6, {'type': 'cell', 'criteria': '>', 'value': 0,'format': deep_green_format})

worksheet.freeze_panes(1, 0)





writerAll.save()
output_excel_file.save()
