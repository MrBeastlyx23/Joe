'''
Generates IBT report -> must create daily inventory files then vLookup the shipdate/BOL
- Delete any random shit at bottom of JSS/30 day inv sheets 
- Ensure Past IBT Columns = ['PO_num', 'X-Dock Notes', 'RD Notes', 'GRN']
''' # (no need to delete for JSS - JSS skipsfooter)

# currently using raw_daily30(macro book) and visibility30 - this is diff than 4/12/21's IBTreport499merge - & 4/13's IBTreport499 (v2 files)

# the 499 COM 7 day is received once/week on Mondays
# --> the 499 7 days will be stacked in raw daily file to form makeshift 30 day after filtering dates >30 days  
# the 490 tracker for 490 B&M 30 day shipped is sent *a few times* per week by Jessi Ledvina

'''
# Log
- sevenDay_499 was added
- Carrier Name column would need to be fixed (there is an empty 2nd line in header that messes it up)
- 4/12/21 the JSS + sevenDay merge was added
- 4/13/21 adding 490 30 day shipped - from 490 tracker (Jessi Ledvina) 
- not adding 490 ready to ship yet
- 4/13/21 adding Past 490 IBT w/ merge
- 4/21/21 490 30 day shipped: sorting by date
- 4/21/21 499 7 day shipped will now be 30 day shippped (stacking 5 reports in raw file then filter >30 days, newest on bottom) variable names will stay the same (sevenDay_499), but excel sheet names will be changed to reflect 30 days
- 5/5/21 switched 490 IBT summ columns to be same as rest of xdocks
- 6/4/21 Added "YOI PO" column after RD Notes - use Live YOI report to find + merge the PO#
- 6/7/21 Added "YOI PO Note" column to the Ready/30day reports that go into visibility report (don't get saved to the inv report folder)
- 06/11/21 changed datatype of supplier to int in 189 JSS (Present). + shipped status + fixed the above from 06/07/21
- 7/22/21 Added code to extract Past IBT Summs directly from pubshare-folder
- 07/28/21 pivoting YOI data when merging w/ 30 day+ready to ship (not needed for IBT summ tab) -0 to prevent duplicate PO's
- 07/28/21 changed 432 ready to ship -- new emails
'''

# Make new folder each YEAR & change filepath at bottom

import numpy as np
from numpy import int64
import pandas as pd
import datetime as dt
from datetime import datetime, date

'''
Read each sheet - present IBT is jss data - skip 7 rows and load PO as str to enable PO_num creation later on

'''
with pd.ExcelFile("C:/Users/mlowisz/Desktop/raw_daily30.xlsm") as file:
    jss_189 = pd.read_excel(file, 'Present_189_IBT', skiprows = 7, dtype = {'PO': str, 'Supplier':int}, skipfooter = 1)
    jss_410 = pd.read_excel(file, 'Present_410_IBT', skiprows = 7, dtype = {'PO': str}, skipfooter = 1)
    jss_429 = pd.read_excel(file, 'Present_429_IBT', skiprows = 7, dtype = {'PO': str}, skipfooter = 1)
    jss_432 = pd.read_excel(file, 'Present_432_IBT', skiprows = 7, dtype = {'PO': str}, skipfooter = 1)
    jss_490 = pd.read_excel(file, 'Present_490_IBT', skiprows = 7, dtype = {'PO': str}, skipfooter = 1)
    jss_499 = pd.read_excel(file, 'Present_499_IBT', skiprows = 7, dtype = {'PO': str}, skipfooter = 1)
    #old_ibt_189 = pd.read_excel(file, 'Past_189_IBT', dtype = {'PO' : str})
    old_ibt_410 = pd.read_excel(file, 'Past_410_IBT', dtype = {'PO' : str})
    #old_ibt_429 = pd.read_excel(file, 'Past_429_IBT', dtype = {'PO' : str})
    #old_ibt_432 = pd.read_excel(file, 'Past_432_IBT', dtype = {'PO' : str})
    #old_ibt_490 = pd.read_excel(file, 'Past_490_IBT', dtype = {'PO' : str}) #added - but diff cols than rest
    #old_ibt_499 = pd.read_excel(file, 'Past_499_IBT', dtype = {'PO' : str})
    thirtyDay_189 = pd.read_excel(file,'189_30day', dtype = {'Order Level2' : str})
    thirtyDay_410 = pd.read_excel(file, '410_30day', dtype = {'batchcode' : str, 'consignee' : str})
    thirtyDay_429 = pd.read_excel(file, '429_30day', dtype = {'Ord Lev2' : str})
    thirtyDay_432 = pd.read_excel(file, '432_30day', dtype = {'Cust Ord Num' : str})
    thirtyDay_490 = pd.read_excel(file, '490_30day', dtype = {'Outbound PO' : str})
    # 7 day 499
    sevenDay_499 = pd.read_excel(file, '499_7day', dtype = {'PO#' : str, 'Ship to Store' : str}) # 4/21/21 now 30 day
    ready_189 = pd.read_excel(file, '189ready')
    ready_410 = pd.read_excel(file, '410ready', dtype = {'production_code' : str, 'batchcode' : str})
    ready_429 = pd.read_excel(file, '429ready', skiprows = 1, skipfooter = 2)
    ready_432 = pd.read_excel(file, '432ready')
    #ready_490 = pd.read_excel(file, '490ready')
    liveYOI = pd.read_excel(file, 'Live_YOI', usecols="A:B")
    buyer = pd.read_excel(file, 'Buyer')
    header = pd.read_excel(file,'header', header = None, index_col = None)

print('Base Raw Files Read')


days_since_ran = 1
# If today is Monday-Sunday (0 through 6) 
if date.today().weekday() == 0: #Monday
    days_since_ran = 3
    days_since_ran490 = 3
elif date.today().weekday() == 2: #Wed
    days_since_ran490 = 2
elif date.today().weekday() == 4: #Fri
    days_since_ran490 = 2
elif date.today().weekday() == 6: #Sun
    days_since_ran = 2
    days_since_ran490 = 2
else: #Tues, Thurs, Sat
    days_since_ran = 1
    days_since_ran490 = 1 #490 is normally ran MWF

yesterdays_date = datetime.now() - dt.timedelta(days=days_since_ran)
yesterdays_date = yesterdays_date.strftime('%m_%d_%y')
yesterdays_date490 = datetime.now() - dt.timedelta(days=days_since_ran490)
yesterdays_date490 = yesterdays_date490.strftime('%m_%d_%y')

print('Extracting Past IBT Summary sheets from pubshare')

yest_189 = "189 IBT Summary Report {}".format(yesterdays_date)
old_ibt_189 = pd.read_excel(r"P:\CROSS DOCK MASTER FILES\Cross Dock Reports\Virtual IBT Reports\189 IBT Cross Dock Open P.O Summary Report 2021.xlsx", yest_189, dtype = {'PO' : int})
yest_429 = "429 IBT Summary Report {}".format(yesterdays_date)
old_ibt_429 = pd.read_excel(r"P:\CROSS DOCK MASTER FILES\Cross Dock Reports\Virtual IBT Reports\429 IBT Cross Dock Open P.O Summary Report 2021.xlsx", yest_429, dtype = {'PO' : int})
yest_432 = "432 IBT Summary Report {}".format(yesterdays_date)
old_ibt_432 = pd.read_excel(r"P:\CROSS DOCK MASTER FILES\Cross Dock Reports\Virtual IBT Reports\432 IBT Cross Dock Open P.O Summary Report 2021.xlsx", yest_432, dtype = {'PO' : int})
yest_490 = "490 IBT Summary Report {}".format(yesterdays_date490)
old_ibt_490 = pd.read_excel(r"P:\CROSS DOCK MASTER FILES\Cross Dock Reports\Virtual IBT Reports\490 IBT Cross Dock Open P.O Summary Report 2021.xlsx", yest_490, dtype = {'PO' : int})
yest_499 = "499 IBT Summary Report {}".format(yesterdays_date)
old_ibt_499 = pd.read_excel(r"P:\CROSS DOCK MASTER FILES\Cross Dock Reports\Virtual IBT Reports\499 IBT Cross Dock Open P.O Summary Report 2021.xlsx", yest_499, dtype = {'PO' : int})

print('Past IBT Summaries Read')


'''
Convert HEADOFFICE to its respective store #. Join store coll w/ PO coll to create PO_num coll
'''
oldIBTcolumns = ['PO_num', 'POWeight', 'X-Dock Notes', 'RD Notes', 'GRN']
oldIBTcolumns2 = ['PO_num', 'POWeight', 'Notes', 'GRN'] # was previously used for 490
old_ibt_189 = old_ibt_189[oldIBTcolumns]
old_ibt_410 = old_ibt_410[oldIBTcolumns]
old_ibt_429 = old_ibt_429[oldIBTcolumns]
old_ibt_432 = old_ibt_432[oldIBTcolumns]
old_ibt_490 = old_ibt_490[oldIBTcolumns]
old_ibt_499 = old_ibt_499[oldIBTcolumns]


listed = [jss_189, jss_410, jss_429, jss_432, jss_490, jss_499]
for dock in listed:
    dock['store'] = dock.Type.str[-4:-1]

jss_189 = jss_189.replace({'store': {'FIC':'189'}})
jss_410 = jss_410.replace({'store': {'FIC':'410'}})
jss_429 = jss_429.replace({'store': {'FIC':'429'}})
jss_432 = jss_432.replace({'store': {'FIC':'432'}})
jss_490 = jss_490.replace({'store': {'FIC':'490'}})
jss_499 = jss_499.replace({'store': {'FIC':'499'}})

listed_jss_Sheet = [jss_189, jss_410, jss_429, jss_432, jss_490, jss_499]
for jss_sheet in listed_jss_Sheet:
    jss_sheet['PO_num'] = jss_sheet['store'] + jss_sheet['PO']
    jss_sheet.PO_num = jss_sheet.PO_num.astype(int)

'''
Merge Buyer List w/ jss data to create category column

'''
jss_189 = pd.merge(jss_189, buyer, how = 'left', copy=True)
jss_410 = pd.merge(jss_410, buyer, how = 'left', copy=True)
jss_429 = pd.merge(jss_429, buyer, how = 'left', copy=True)
jss_432 = pd.merge(jss_432, buyer, how = 'left', copy=True)
jss_490 = pd.merge(jss_490, buyer, how = 'left', copy=True)
jss_499 = pd.merge(jss_499, buyer, how = 'left', copy=True)

jss_189['daysAtDock'] = (dt.datetime.now() - jss_189['Delivery Date']).dt.days
jss_410['daysAtDock'] = (dt.datetime.now() - jss_410['Delivery Date']).dt.days
jss_429['daysAtDock'] = (dt.datetime.now() - jss_429['Delivery Date']).dt.days
jss_432['daysAtDock'] = (dt.datetime.now() - jss_432['Delivery Date']).dt.days
jss_490['daysAtDock'] = (dt.datetime.now() - jss_490['Delivery Date']).dt.days
jss_499['daysAtDock'] = (dt.datetime.now() - jss_499['Delivery Date']).dt.days

JSS_final_rest1_189 = jss_189[(jss_189['store'] == '189') & (jss_189['Category'] != 'PRODUCE')]
JSS_final_produce1_189 = jss_189[(jss_189['store'] == '189') & (jss_189['Category'] == 'PRODUCE')]
JSS_IBT_Produce1_189 = jss_189[(jss_189['store'] != '189') & (jss_189['Category'] == 'PRODUCE')]
JSS_IBT_rest1_189 = jss_189[(jss_189['store'] != '189') & (jss_189['Category'] != 'PRODUCE')]

JSS_final_rest1_189 = JSS_final_rest1_189.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_IBT_Produce1_189 = JSS_IBT_Produce1_189.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_IBT_rest1_189 = JSS_IBT_rest1_189.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_final_produce1_189 = JSS_final_produce1_189.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)

final_IBT1_189 = JSS_final_rest1_189.append(header)
final_IBT2_189 = final_IBT1_189.append(JSS_final_produce1_189)
final_IBT3_189 = final_IBT2_189.append(header)
final_IBT4_189 = final_IBT3_189.append(JSS_IBT_rest1_189)
final_IBT5_189 = final_IBT4_189.append(header)
final_JSS_189 = final_IBT5_189.append(JSS_IBT_Produce1_189)

final_JSS_189 = pd.merge(final_JSS_189, old_ibt_189, how = 'left', on = ['PO_num', 'POWeight'])


JSS_final_rest1_410 = jss_410[(jss_410['store'] == '410') & (jss_410['Category'] != 'PRODUCE')]
JSS_final_produce1_410 = jss_410[(jss_410['store'] == '410') & (jss_410['Category'] == 'PRODUCE')]
JSS_IBT_Produce1_410 = jss_410[(jss_410['store'] != '410') & (jss_410['Category'] == 'PRODUCE')]
JSS_IBT_rest1_410 = jss_410[(jss_410['store'] != '410') & (jss_410['Category'] != 'PRODUCE')]

JSS_final_rest1_410 = JSS_final_rest1_410.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_IBT_Produce1_410 = JSS_IBT_Produce1_410.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_IBT_rest1_410 = JSS_IBT_rest1_410.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_final_produce1_410 = JSS_final_produce1_410.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)

final_IBT1_410 = JSS_final_rest1_410.append(header)
final_IBT2_410 = final_IBT1_410.append(JSS_final_produce1_410)
final_IBT3_410 = final_IBT2_410.append(header)
final_IBT4_410 = final_IBT3_410.append(JSS_IBT_rest1_410)
final_IBT5_410 = final_IBT4_410.append(header)
final_JSS_410 = final_IBT5_410.append(JSS_IBT_Produce1_410)

final_JSS_410 = pd.merge(final_JSS_410, old_ibt_410, how = 'left', on = ['PO_num', 'POWeight'])


JSS_final_rest1_429 = jss_429[(jss_429['store'] == '429') & (jss_429['Category'] != 'PRODUCE')]
JSS_final_produce1_429 = jss_429[(jss_429['store'] == '429') & (jss_429['Category'] == 'PRODUCE')]
JSS_IBT_Produce1_429 = jss_429[(jss_429['store'] != '429') & (jss_429['Category'] == 'PRODUCE')]
JSS_IBT_rest1_429 = jss_429[(jss_429['store'] != '429') & (jss_429['Category'] != 'PRODUCE')]

JSS_final_rest1_429 = JSS_final_rest1_429.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_IBT_Produce1_429 = JSS_IBT_Produce1_429.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_IBT_rest1_429 = JSS_IBT_rest1_429.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_final_produce1_429 = JSS_final_produce1_429.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)

final_IBT1_429 = JSS_final_rest1_429.append(header)
final_IBT2_429 = final_IBT1_429.append(JSS_final_produce1_429)
final_IBT3_429 = final_IBT2_429.append(header)
final_IBT4_429 = final_IBT3_429.append(JSS_IBT_rest1_429)
final_IBT5_429 = final_IBT4_429.append(header)
final_JSS_429 = final_IBT5_429.append(JSS_IBT_Produce1_429)

final_JSS_429 = pd.merge(final_JSS_429, old_ibt_429, how = 'left', on = ['PO_num', 'POWeight'])


JSS_final_rest1_432 = jss_432[(jss_432['store'] == '432') & (jss_432['Category'] != 'PRODUCE')]
JSS_final_produce1_432 = jss_432[(jss_432['store'] == '432') & (jss_432['Category'] == 'PRODUCE')]
JSS_IBT_Produce1_432 = jss_432[(jss_432['store'] != '432') & (jss_432['Category'] == 'PRODUCE')]
JSS_IBT_rest1_432 = jss_432[(jss_432['store'] != '432') & (jss_432['Category'] != 'PRODUCE')]

JSS_final_rest1_432 = JSS_final_rest1_432.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_IBT_Produce1_432 = JSS_IBT_Produce1_432.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_IBT_rest1_432 = JSS_IBT_rest1_432.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_final_produce1_432 = JSS_final_produce1_432.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)

final_IBT1_432 = JSS_final_rest1_432.append(header)
final_IBT2_432 = final_IBT1_432.append(JSS_final_produce1_432)
final_IBT3_432 = final_IBT2_432.append(header)
final_IBT4_432 = final_IBT3_432.append(JSS_IBT_rest1_432)
final_IBT5_432 = final_IBT4_432.append(header)
final_JSS_432 = final_IBT5_432.append(JSS_IBT_Produce1_432)

final_JSS_432 = pd.merge(final_JSS_432, old_ibt_432, how = 'left', on = ['PO_num', 'POWeight'])


JSS_final_rest1_490 = jss_490[(jss_490['store'] == '490') & (jss_490['Category'] != 'PRODUCE')]
JSS_final_produce1_490 = jss_490[(jss_490['store'] == '490') & (jss_490['Category'] == 'PRODUCE')]
JSS_IBT_Produce1_490 = jss_490[(jss_490['store'] != '490') & (jss_490['Category'] == 'PRODUCE')]
JSS_IBT_rest1_490 = jss_490[(jss_490['store'] != '490') & (jss_490['Category'] != 'PRODUCE')]

JSS_final_rest1_490 = JSS_final_rest1_490.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_IBT_Produce1_490 = JSS_IBT_Produce1_490.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_IBT_rest1_490 = JSS_IBT_rest1_490.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_final_produce1_490 = JSS_final_produce1_490.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)

final_IBT1_490 = JSS_final_rest1_490.append(header)
final_IBT2_490 = final_IBT1_490.append(JSS_final_produce1_490)
final_IBT3_490 = final_IBT2_490.append(header)
final_IBT4_490 = final_IBT3_490.append(JSS_IBT_rest1_490)
final_IBT5_490 = final_IBT4_490.append(header)
final_JSS_490 = final_IBT5_490.append(JSS_IBT_Produce1_490)

final_JSS_490 = pd.merge(final_JSS_490, old_ibt_490, how = 'left', on = ['PO_num', 'POWeight'])


JSS_final_rest1_499 = jss_499[(jss_499['store'] == '499') & (jss_499['Category'] != 'PRODUCE')]
JSS_final_produce1_499 = jss_499[(jss_499['store'] == '499') & (jss_499['Category'] == 'PRODUCE')]
JSS_IBT_Produce1_499 = jss_499[(jss_499['store'] != '499') & (jss_499['Category'] == 'PRODUCE')]
JSS_IBT_rest1_499 = jss_499[(jss_499['store'] != '499') & (jss_499['Category'] != 'PRODUCE')]

JSS_final_rest1_499 = JSS_final_rest1_499.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_IBT_Produce1_499 = JSS_IBT_Produce1_499.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_IBT_rest1_499 = JSS_IBT_rest1_499.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
JSS_final_produce1_499 = JSS_final_produce1_499.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)

final_IBT1_499 = JSS_final_rest1_499.append(header)
final_IBT2_499 = final_IBT1_499.append(JSS_final_produce1_499)
final_IBT3_499 = final_IBT2_499.append(header)
final_IBT4_499 = final_IBT3_499.append(JSS_IBT_rest1_499)
final_IBT5_499 = final_IBT4_499.append(header)
final_JSS_499 = final_IBT5_499.append(JSS_IBT_Produce1_499)

final_JSS_499 = pd.merge(final_JSS_499, old_ibt_499, how = 'left', on = ['PO_num', 'POWeight'])


## 6/4/21 - merge of YOI PO - no 410
final_JSS_189 = pd.merge(final_JSS_189, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])
final_JSS_429 = pd.merge(final_JSS_429, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])
final_JSS_432 = pd.merge(final_JSS_432, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])
final_JSS_490 = pd.merge(final_JSS_490, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])
final_JSS_499 = pd.merge(final_JSS_499, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])


IBT_columns1 = ['PO_num', 'Supplier', 'SupplierName', 'POWeight', 'Order Value', 'Orderdate', 'Delivery Date', 'daysAtDock', 
                'Category', 'X-Dock Notes', 'RD Notes', 'YOI PO Note', 'GRN', 'Buyer', 'ShippingType']
IBT_columns2 = ['PO_num', 'Supplier', 'SupplierName', 'POWeight', 'Order Value', 'Orderdate', 'Delivery Date', 'daysAtDock', 
                'Category', 'X-Dock Notes', 'RD Notes', 'GRN', 'Buyer', 'ShippingType']
#4/13/21 added 'Notes' + 'GRN' to cols2 #5/5/21 switched to regular columns

final_JSS_189 = final_JSS_189[IBT_columns1]
final_JSS_410 = final_JSS_410[IBT_columns2]
final_JSS_429 = final_JSS_429[IBT_columns1] 
final_JSS_432 = final_JSS_432[IBT_columns1]
final_JSS_490 = final_JSS_490[IBT_columns1]
final_JSS_499 = final_JSS_499[IBT_columns1]
print('JSS Sheets Created')

'''
Delete any random bullshit (non-numeric) in PO columns
'''
thirtyDay_189['Order Level2'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'')
thirtyDay_429['Ord Lev2'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'')
thirtyDay_432['Cust Ord Num'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'')
thirtyDay_490['Outbound PO'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'') ###
sevenDay_499['PO#'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'') ######### 7 day 499 #####


'''
Drop N/A and combine store/PO colls for 410, then rename consignee col. to Store
'''
thirtyDay_410.dropna(axis=0, how='any', inplace = True, subset = ['consignee', 'batchcode'])

thirtyDay_410['PO_num'] = thirtyDay_410['consignee'] + thirtyDay_410['batchcode']
thirtyDay_410['PO_num'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'')

thirtyDay_410 = thirtyDay_410.rename(columns = {'consignee' : 'Store', 'ship_date':'Ship Date', 'pronumber':'BOL#'})

'''
Rename PO_num columns in 189, 429, 432 and create store # columns
'''
thirtyDay_189 = thirtyDay_189.rename(columns = {'Order Level2' : 'PO_num', 'Confirm Date' : 'Ship Date', 'Order Num': 'BOL#'})
thirtyDay_429 = thirtyDay_429.rename(columns = {'Ord Lev2' : 'PO_num', 'Doc Num': 'BOL#', 'Ord To Ship Date':'Ship Date'})
thirtyDay_432 = thirtyDay_432.rename(columns = {'Cust Ord Num' : 'PO_num', 'Confirm Date':'Ship Date', 'Load Num':'BOL#'})
thirtyDay_490 = thirtyDay_490.rename(columns = {'Outbound PO' : 'PO_num', 'BM #':'BOL#'}) ##

thirtyDay_189['Store'] = thirtyDay_189['PO_num'].str[:3]
thirtyDay_429['Store'] = thirtyDay_429['PO_num'].str[:3]
thirtyDay_432['Store'] = thirtyDay_432['PO_num'].str[:3]
thirtyDay_490['Store'] = thirtyDay_490['PO_num'].str[:3] ##

'''
Drop N/A then convert store and PO_num to int
Arrange final 30 day inv columns
'''
thirty_189 = ['Store', 'PO_num', 'BOL#', 'Ship Date', 'Customer', 'Order Level3']

thirtyDay_189.dropna(axis=0, how='any', inplace = True, subset = ['PO_num', 'Store'])
thirtyDay_189.PO_num = thirtyDay_189.PO_num.astype(int)
thirtyDay_189.Store = thirtyDay_189.Store.astype(int)
thirtyDay_189_final = thirtyDay_189[thirty_189]

thirty_410 = ['Store', 'PO_num', 'BOL#', 'Ship Date', 'qty', 'gross_wt', 'item_num']

thirtyDay_410.dropna(axis=0, how='any', inplace = True, subset = ['PO_num', 'Store'])
thirtyDay_410.PO_num = thirtyDay_410.PO_num.astype(int64)
thirtyDay_410.Store = thirtyDay_410.Store.astype(int)
thirtyDay_410_final = thirtyDay_410[thirty_410]

thirty_429 = ['Store', 'PO_num', 'BOL#', 'Ship Date','Cust Code', 'Ord Lev1', 'Ord Lev3', 'Ord Ship Qty']

thirtyDay_429.dropna(axis=0, how='any', inplace = True, subset = ['PO_num', 'Store'])
thirtyDay_429.PO_num = thirtyDay_429.PO_num.astype(int)
thirtyDay_429.Store = thirtyDay_429.Store.astype(int)
thirtyDay_429_final = thirtyDay_429[thirty_429]

thirty_432 = ['Store', 'PO_num', 'BOL#', 'Ship Date', 'Ord Ship Qty', 'Total Wgt Net']
thirtyDay_432.dropna(axis=0, how='any', inplace = True, subset = ['PO_num', 'Store'])
thirtyDay_432.PO_num = thirtyDay_432.PO_num.astype(int64)
thirtyDay_432.Store = thirtyDay_432.Store.astype(int)
thirtyDay_432_final = thirtyDay_432[thirty_432]

# we are only interested in rows with <=30 days since ship date
# we first convert string to date
thirtyDay_490 = thirtyDay_490[thirtyDay_490['Ship Date'] != 'TBD'] # also need to delete these TBD that are added. # it would be smarter to make more exeptions though
thirtyDay_490 = thirtyDay_490[thirtyDay_490['BOL#'] != 'pushed'] # BM# is already renamed to BOL#
thirtyDay_490['Ship Date'] = pd.to_datetime(thirtyDay_490['Ship Date'])
thirtyDay_490['daysCounter'] = (dt.datetime.now() - thirtyDay_490['Ship Date']).dt.days #dtype = days
thirtyDay_490 = thirtyDay_490[thirtyDay_490['daysCounter'] < 31] ##
thirty_490 = ['Store', 'PO_num', 'BOL#', 'Ship Date', 'Vendor']
thirtyDay_490.dropna(axis=0, how='any', inplace = True, subset = ['PO_num', 'Store'])
thirtyDay_490.PO_num = thirtyDay_490.PO_num.astype(int)
thirtyDay_490.Store = thirtyDay_490.Store.astype(int)
thirtyDay_490_final = thirtyDay_490[thirty_490]
# 4/21/21 sorting by ship date
thirtyDay_490_final = thirtyDay_490_final.sort_values(by = ['Ship Date'] , inplace = False, ascending = True)


# 4/21/21 update of making 7 day into a 30 day (stacked)
#thirtyDay_490['Ship Date'] = pd.to_datetime(thirtyDay_490['Ship Date'])
#thirtyDay_490['Ship Date'] = pd.to_datetime(thirtyDay_490['Ship Date'], format='%d/%m/%Y')
#thirtyDay_490['Ship Date'] = datetime.strptime(thirtyDay_490['Ship Date'], '%d/%m/%Y').date() # dtype = date
sevenDay_499['daysCounter'] = (dt.datetime.now() - sevenDay_499['Ship Date']).dt.days #dtype = days
sevenDay_499 = sevenDay_499[sevenDay_499['daysCounter'] < 31] ##
'''
 499 Drop N/A and Combine store & PO# (first 5 chars) cols to create PO_num col for 499
'''
sevenDay_499.dropna(axis=0, how='any', inplace = True, subset = ['PO#', 'Ship to Store'])
sevenDay_499['PO_num'] = sevenDay_499['Ship to Store'] + sevenDay_499['PO#'].str[:5]
## ***** Load ID may not be BOL# **
sevenDay_499 = sevenDay_499.rename(columns = {'Ship to Store' : 'Store', 'Load ID': 'BOL#'})
# Arrange final 7 day columns - does not include 'Carrier Name'
seven_499 = ['Store', 'PO_num', 'BOL#', 'Ship Date', 'Shipped Quantity', 'Cases', 'PO Weight']
# convert store and PO_num to int
sevenDay_499.PO_num = sevenDay_499.PO_num.astype(int64)
sevenDay_499.Store = sevenDay_499.Store.astype(int)
sevenDay_499_final = sevenDay_499[seven_499]


thirtyDay_189_vl = thirtyDay_189_final[['PO_num', 'BOL#', 'Ship Date']]
test_189_jss = pd.merge(final_JSS_189, thirtyDay_189_vl, how = 'inner', on='PO_num', copy = False)
test_189_jss = test_189_jss[test_189_jss['Orderdate'] <= test_189_jss['Ship Date']]
test_189_jss = test_189_jss[['PO_num', 'BOL#', 'Ship Date']]
final_JSS_189 = pd.merge(final_JSS_189, test_189_jss, how='left', on='PO_num', copy = False)
final_JSS_189 = final_JSS_189.drop_duplicates(subset= 'PO_num')

thirtyDay_410_vl = thirtyDay_410_final[['PO_num', 'BOL#', 'Ship Date']]
test_410_jss = pd.merge(final_JSS_410, thirtyDay_410_vl, how = 'inner', on='PO_num')
test_410_jss = test_410_jss[test_410_jss['Orderdate'] <= test_410_jss['Ship Date']]
test_410_jss = test_410_jss[['PO_num', 'BOL#', 'Ship Date']]
final_JSS_410 = pd.merge(final_JSS_410, test_410_jss, how='left', on='PO_num', copy = False)
final_JSS_410 = final_JSS_410.drop_duplicates(subset= 'PO_num')

thirtyDay_429_vl = thirtyDay_429_final[['PO_num', 'BOL#', 'Ship Date']]
test_429_jss = pd.merge(final_JSS_429, thirtyDay_429_vl, how = 'inner', on='PO_num', copy = False)
test_429_jss = test_429_jss[test_429_jss['Orderdate'] <= test_429_jss['Ship Date']]
test_429_jss = test_429_jss[['PO_num', 'BOL#', 'Ship Date']]
final_JSS_429 = pd.merge(final_JSS_429, test_429_jss, how='left', on='PO_num', copy = False)
final_JSS_429 = final_JSS_429.drop_duplicates(subset= 'PO_num')

thirtyDay_432_vl = thirtyDay_432_final[['PO_num', 'BOL#', 'Ship Date']]
test_432_jss = pd.merge(final_JSS_432, thirtyDay_432_vl, how = 'inner', on='PO_num', copy = False)
test_432_jss = test_432_jss[test_432_jss['Orderdate'] <= test_432_jss['Ship Date']]
test_432_jss = test_432_jss[['PO_num', 'BOL#', 'Ship Date']]
final_JSS_432 = pd.merge(final_JSS_432, test_432_jss, how='left', on='PO_num', copy = False)
final_JSS_432 = final_JSS_432.drop_duplicates(subset= 'PO_num')

# 4/13/21
thirtyDay_490_vl = thirtyDay_490_final[['PO_num', 'BOL#', 'Ship Date']]
test_490_jss = pd.merge(final_JSS_490, thirtyDay_490_vl, how = 'inner', on='PO_num', copy = False)
test_490_jss = test_490_jss[test_490_jss['Orderdate'] <= test_490_jss['Ship Date']]
test_490_jss = test_490_jss[['PO_num', 'BOL#', 'Ship Date']]
final_JSS_490 = pd.merge(final_JSS_490, test_490_jss, how='left', on='PO_num', copy = False)
final_JSS_490 = final_JSS_490.drop_duplicates(subset= 'PO_num')

#new merge implemented on 4/12/21
sevenDay_499_vl = sevenDay_499_final[['PO_num', 'BOL#', 'Ship Date']]
test_499_jss = pd.merge(final_JSS_499, sevenDay_499_vl, how = 'inner', on='PO_num', copy = False)
test_499_jss = test_499_jss[test_499_jss['Orderdate'] <= test_499_jss['Ship Date']]
test_499_jss = test_499_jss[['PO_num', 'BOL#', 'Ship Date']]
final_JSS_499 = pd.merge(final_JSS_499, test_499_jss, how='left', on='PO_num', copy = False)
final_JSS_499 = final_JSS_499.drop_duplicates(subset= 'PO_num')

print('30 day+JSS merged')


'''
Clean all PO# columns for ready to ship
'''
ready_189['Invt Lev2'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'')
ready_429['Level 2'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'')
#ready_432['L2L3 PO'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'')
ready_432['Cust Ord Num'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'')
ready_410['batchcode'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'')
ready_410['production_code'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'')

'''
Rename/Create PO_num column for ready to ship
'''
ready_410['PO_num'] = ready_410['production_code'] + ready_410['batchcode']
ready_189 = ready_189.rename(columns = {'Invt Lev2' : 'PO_num'})
ready_429 = ready_429.rename(columns = {'Level 2' : 'PO_num'})
#ready_432 = ready_432.rename(columns = {'L2L3 PO' : 'PO_num'})
ready_432 = ready_432.rename(columns = {'Cust Ord Num' : 'PO_num'})

'''
Convert datatype of PO columns to make pivot table work for ready to ship
'''
ready_189.PO_num = ready_189.PO_num.astype(str)
ready_429.PO_num = ready_429.PO_num.astype(str)
ready_432.PO_num = ready_432.PO_num.astype(str)

'''
Create Store columns for ready to ship
'''
ready_410 = ready_410.rename(columns = {'production_code' : 'Store'})
ready_189['Store'] = ready_189['PO_num'].str[:3]
ready_429['Store'] = ready_429['PO_num'].str[:3]
ready_432['Store'] = ready_432['PO_num'].str[:3]


# convert the date+time to just a date -- for pivot table
ready_432['Ord To Ship Date'] = pd.to_datetime(ready_432['Ord To Ship Date']) #, format='%m/%d/%Y'


'''
Create Pivot Tables for Ready To Ship
'''
pivot189 = pd.pivot_table(ready_189, index = ['Store', 'PO_num','Cust Code', 'Invt Lev3', 'Invt Recd Date'],  values = 'On Hand Wgt Net', aggfunc = np.sum)
pivot189 = pivot189.reset_index()

pivot429 = pd.pivot_table(ready_429, index = ['Store', 'PO_num','Level 3', 'Storage_Type', 'Invt Org Recd Date'],  values = 'On Hand Wgt', aggfunc = np.sum)
pivot429 = pivot429.reset_index()

#pivot432 = pd.pivot_table(ready_432, index = ['Store', 'PO_num','Item Desc', 'From Rcpt Ship Name', 'Invt Org Recd Date'],  values = 'On Hand Wgt', aggfunc = np.sum)
pivot432 = pd.pivot_table(ready_432, index = ['Store', 'PO_num', 'Ord Num', 'Ord Date', 'Ord To Ship Date', 'Cust Code'],  values = 'Ship Wgt', aggfunc = np.sum)
pivot432 = pivot432.reset_index() 

ready410 = ['Store', 'PO_num', 'job_num', 'item_num', 'recv_date', 'qty_onhand', 'net_wt']
ready_410 = ready_410[ready410]

'''
add days at dock columns in ready to ship
'''
ready_410['recv_date'] = pd.to_datetime(ready_410['recv_date'])
pivot189['daysAtDock'] = (dt.datetime.now() - pivot189['Invt Recd Date']).dt.days
#pivot432['daysAtDock'] = (dt.datetime.now() - pivot432['Invt Org Recd Date']).dt.days
pivot432['daysAtDock'] = (dt.datetime.now() - pivot432['Ord Date']).dt.days
pivot429['daysAtDock'] = (dt.datetime.now() - pivot429['Invt Org Recd Date']).dt.days
ready_410['daysAtDock'] = (dt.datetime.now() - ready_410['recv_date']).dt.days

'''
Sort ready to ship days at dock columns
'''
pivot189 = pivot189.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
ready_410 = ready_410.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
pivot429 = pivot429.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)
pivot432 = pivot432.sort_values(by = ['daysAtDock'] , inplace = False, ascending = False)

'''
Add notes to ready to ship DFs
'''
pivot189['Notes'] = np.where(pivot189['daysAtDock'] < 11, "Heading out on next available truck", "AV in motion")
pivot429['Notes'] = np.where(pivot429['daysAtDock'] < 11, "Heading out on next available truck", "AV in motion")
pivot432['Notes'] = np.where(pivot432['daysAtDock'] < 11, "Heading out on next available truck", "AV in motion")
pivot432['Shipped Status'] = 'Shipping ' + pivot432['Ord To Ship Date'].dt.strftime('%m-%d-%y') + ' under BOL# ' + pivot432['Ord Num'].astype(str)
ready_410['Notes'] = np.where(ready_410['daysAtDock'] < 11, "Heading out on next available truck", "AV in motion")

'''
Convert dtypes
'''
pivot189.PO_num = pivot189.PO_num.astype(int)
pivot189.Store = pivot189.Store.astype(int)
ready_410.PO_num = ready_410.PO_num.astype(int)
ready_410.Store = ready_410.Store.astype(int)
pivot429.PO_num = pivot429.PO_num.astype(int)
pivot429.Store = pivot429.Store.astype(int)
pivot432.PO_num = pivot432.PO_num.astype(int)
pivot432.Store = pivot432.Store.astype(int)
# ready to ship done



# 06/7/21 and 06/11/21
### Tabs for OutPut tabs (visibility file)
'''
These are not saved to the DailyInv/30DayShip folder - they are only for the 
### Differences
# PO Shipped Status - combining DateShipped + BOL#
# YOI Note (or YOI Identifier) col for ReadyToShip (or 30day) - finds that day's Live YOI PO's and returns that generic msg
'''
# before merging, we remove dupe POs #added 07/28/21
liveYOI.drop_duplicates(subset ="YOI PO", keep = False, inplace = True)
vis_thirtyDay_189_final = pd.merge(thirtyDay_189_final, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])
vis_thirtyDay_429_final = pd.merge(thirtyDay_429_final, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])
vis_thirtyDay_432_final = pd.merge(thirtyDay_432_final, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])
vis_thirtyDay_490_final = pd.merge(thirtyDay_490_final, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])
vis_sevenDay_499_final = pd.merge(sevenDay_499_final, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])
vis_pivot189 = pd.merge(pivot189, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])[['Store','PO_num','Cust Code','Invt Lev3','Invt Recd Date','On Hand Wgt Net','daysAtDock','Notes','YOI PO Note']]
vis_pivot429 = pd.merge(pivot429, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])[['Store','PO_num','Level 3','Storage_Type','Invt Org Recd Date','On Hand Wgt','daysAtDock','Notes','YOI PO Note']]
#vis_pivot432 = pd.merge(pivot432, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])[['Store','PO_num','Item Desc','From Rcpt Ship Name','Invt Org Recd Date','On Hand Wgt','daysAtDock','Notes','YOI PO Note']]
vis_pivot432 = pd.merge(pivot432, liveYOI, how = 'left', left_on = ['PO_num'], right_on = ['YOI PO'])[['Store','PO_num','Ord Num','Ord Date','Ord To Ship Date','Cust Code','Ship Wgt','daysAtDock','Notes', 'Shipped Status', 'YOI PO Note']]
# rename the YOI Note column for 30daySheet
vis_thirtyDay_189_final.rename(columns={'YOI PO Note': 'YOI IDENTIFIER'}, inplace = True)
vis_thirtyDay_429_final.rename(columns={'YOI PO Note': 'YOI IDENTIFIER'}, inplace = True)
vis_thirtyDay_432_final.rename(columns={'YOI PO Note': 'YOI IDENTIFIER'}, inplace = True)
vis_thirtyDay_490_final.rename(columns={'YOI PO Note': 'YOI IDENTIFIER'}, inplace = True)
vis_sevenDay_499_final.rename(columns={'YOI PO Note': 'YOI IDENTIFIER'}, inplace = True)
# create new col - combining Ship date and BOL#
vis_thirtyDay_189_final['Shipped Status'] = 'Shipped ' + vis_thirtyDay_189_final['Ship Date'].dt.strftime('%m-%d-%y') + ' under BOL# ' + vis_thirtyDay_189_final['BOL#'].astype(str)
vis_thirtyDay_429_final['Shipped Status'] = 'Shipped ' + vis_thirtyDay_429_final['Ship Date'].dt.strftime('%m-%d-%y') + ' under BOL# ' + vis_thirtyDay_429_final['BOL#'].astype(str)
vis_thirtyDay_432_final['Shipped Status'] = 'Shipped ' + vis_thirtyDay_432_final['Ship Date'].dt.strftime('%m-%d-%y') + ' under BOL# ' + vis_thirtyDay_432_final['BOL#'].astype(str)
vis_thirtyDay_490_final['Shipped Status'] = 'Shipped ' + vis_thirtyDay_490_final['Ship Date'].dt.strftime('%m-%d-%y') + ' under BOL# ' + vis_thirtyDay_490_final['BOL#'].astype(str)
vis_sevenDay_499_final['Shipped Status'] = 'Shipped ' + vis_sevenDay_499_final['Ship Date'].dt.strftime('%m-%d-%y') + ' under BOL# ' + vis_sevenDay_499_final['BOL#'].astype(str)
# final columns to be used for output 'visibility'
vis_thirtyDay_189_final = vis_thirtyDay_189_final[['Store','PO_num','BOL#','Ship Date','Customer','Order Level3','Shipped Status','YOI IDENTIFIER']]
vis_thirtyDay_429_final = vis_thirtyDay_429_final[['Store','PO_num','BOL#','Ship Date','Cust Code','Ord Lev1','Ord Lev3','Ord Ship Qty','Shipped Status','YOI IDENTIFIER']]
vis_thirtyDay_432_final = vis_thirtyDay_432_final[['Store','PO_num','BOL#','Ship Date','Ord Ship Qty','Total Wgt Net','Shipped Status','YOI IDENTIFIER']]
vis_thirtyDay_490_final = vis_thirtyDay_490_final[['Store','PO_num','BOL#','Ship Date','Vendor','Shipped Status','YOI IDENTIFIER']]
vis_sevenDay_499_final = vis_sevenDay_499_final[['Store','PO_num','BOL#','Ship Date','Shipped Quantity','Cases','PO Weight','Shipped Status','YOI IDENTIFIER']]


'''
Saving to visibility
'''
 # will add date to sheet name
berk_ready = "189 Ready To Ship {}".format(datetime.now().strftime('%m_%d_%y'))
berk_thirty = "189 30 Day Inv {}".format(datetime.now().strftime('%m_%d_%y'))
cwi_rede = "429 Ready To Ship {}".format(datetime.now().strftime('%m_%d_%y'))
cwi_thirty = "429 30 Day Inv {}".format(datetime.now().strftime('%m_%d_%y'))
ccs_rede = "432 Ready To Ship {}".format(datetime.now().strftime('%m_%d_%y'))
ccs_thirty = "432 30 Day Inv {}".format(datetime.now().strftime('%m_%d_%y'))
bm_thirty = "490 30 Day Inv {}".format(datetime.now().strftime('%m_%d_%y')) ###
com_seven = "499 30 Day Inv {}".format(datetime.now().strftime('%m_%d_%y')) ## 4/21/21 changed from 7 to 30
berk_IBT = "189 IBT Summary Report {}".format(datetime.now().strftime('%m_%d_%y'))
cwi_IBT = "429 IBT Summary Report {}".format(datetime.now().strftime('%m_%d_%y'))
ccs_IBT = "432 IBT Summary Report {}".format(datetime.now().strftime('%m_%d_%y'))
com_IBT = "499 IBT Summary Report {}".format(datetime.now().strftime('%m_%d_%y'))
bm_IBT = "490 IBT Summary Report {}".format(datetime.now().strftime('%m_%d_%y')) ###

 # will add date to sheet name
with pd.ExcelWriter("C:/Users/mlowisz/Desktop/visibility30.xlsx", date_format='MM/DD/YYYY', datetime_format='MM/DD/YYYY') as writer:
    vis_pivot189.to_excel(writer, sheet_name = berk_ready, index = False)
    vis_thirtyDay_189_final.to_excel(writer, sheet_name = berk_thirty, index = False)
    final_JSS_189.to_excel(writer, sheet_name = berk_IBT, index = False)
    ready_410.to_excel(writer, sheet_name = '410 Ready To Ship', index = False)
    thirtyDay_410_final.to_excel(writer, sheet_name = '410 30 Day Inv', index = False)
    final_JSS_410.to_excel(writer, sheet_name = '410 IBT Summary Report', index = False)
    vis_pivot429.to_excel(writer, sheet_name = cwi_rede, index = False)
    vis_thirtyDay_429_final.to_excel(writer, sheet_name = cwi_thirty, index = False)
    final_JSS_429.to_excel(writer, sheet_name = cwi_IBT, index = False)
    vis_pivot432.to_excel(writer, sheet_name = ccs_rede, index = False)
    vis_thirtyDay_432_final.to_excel(writer, sheet_name = ccs_thirty, index = False)
    final_JSS_432.to_excel(writer, sheet_name = ccs_IBT, index = False)
    vis_sevenDay_499_final.to_excel(writer, sheet_name = com_seven, index = False) # 499 - only 1 per week
    #sevenDay_499_final.to_excel(writer, sheet_name = '499 7 Day Shipped', index = False) # old way w/o date
    final_JSS_499.to_excel(writer, sheet_name = com_IBT, index = False)
    vis_thirtyDay_490_final.to_excel(writer, sheet_name = bm_thirty, index = False)
    final_JSS_490.to_excel(writer, sheet_name = bm_IBT, index = False)

    
print("Visibility File Saved On Desktop - visibility30.xlsx (non-macro book)")


berkpath = r"P:\CROSS DOCK MASTER FILES\Cross Dock Inventory\Daily Xdock Virtual Inventories\Ready to ship Daily Inventories\189 Daily Inventory 2021\189  Daily Inventory and 30 Day Ship {}.xlsx".format(datetime.now().strftime('%m.%d.%y'))
with pd.ExcelWriter(berkpath) as writer:
	pivot189.to_excel(writer, sheet_name = '189 Ready To Ship', index = False)
	thirtyDay_189_final.to_excel(writer, sheet_name = '189 30 Day Inv', index = False)

'''
hennepath = r"P:\CROSS DOCK MASTER FILES\Cross Dock Inventory\Daily Xdock Virtual Inventories\410 Daily Inventory 2020\410 Daily Inventory and 30 Day Ship {}.xlsx".format(datetime.now().strftime('%m.%d.%y'))
with pd.ExcelWriter(hennepath) as writer:
	ready_410.to_excel(writer, sheet_name = '410 Ready To Ship', index = False)
	thirtyDay_410_final.to_excel(writer, sheet_name = '410 30 Day Inv', index = False)
'''

CWIpath = r"P:\CROSS DOCK MASTER FILES\Cross Dock Inventory\Daily Xdock Virtual Inventories\Ready to ship Daily Inventories\429 Daily Inventory 2021\429 Daily Inventory and 30 Day Ship {}.xlsx".format(datetime.now().strftime('%m.%d.%y'))
with pd.ExcelWriter(CWIpath) as writer:
	pivot429.to_excel(writer, sheet_name = '429 Ready To Ship', index = False)
	thirtyDay_429_final.to_excel(writer, sheet_name = '429 30 Day Inv', index = False)

CLFpath = r"P:\CROSS DOCK MASTER FILES\Cross Dock Inventory\Daily Xdock Virtual Inventories\Ready to ship Daily Inventories\432 Daily Inventory 2021\432 Daily Inventory and 30 Day Ship {}.xlsx".format(datetime.now().strftime('%m.%d.%y'))
with pd.ExcelWriter(CLFpath) as writer:
	pivot432.to_excel(writer, sheet_name = '432 Ready To Ship', index = False)
	thirtyDay_432_final.to_excel(writer, sheet_name = '432 30 Day Inv', index = False)

# Since we get sent a 7 day shipped report each Monday, this currently saves duplicated files throughout the week
# only one will be needed - can delete duplicates
COMpath = r"P:\CROSS DOCK MASTER FILES\Cross Dock Inventory\Daily Xdock Virtual Inventories\Ready to ship Daily Inventories\499 30 Day Shipped Inventory 2021\499 30 Day Ship {}.xlsx".format(datetime.now().strftime('%m.%d.%y')) # 4/21/21 changed 7 to 30. saved in same folder
with pd.ExcelWriter(COMpath) as writer:
	sevenDay_499_final.to_excel(writer, sheet_name = '499 30 Day Inv', index = False) # 4/21/21 changed 7 to 30

# The 490 tracker file is sent irregularly throughout the week, --> this currently saves duplicated files throughout the week
# only one will be needed - can delete duplicates
BMpath = r"P:\CROSS DOCK MASTER FILES\Cross Dock Inventory\Daily Xdock Virtual Inventories\Ready to ship Daily Inventories\490 30 Day Shipped Inventory 2021\490 30 Day Ship {}.xlsx".format(datetime.now().strftime('%m.%d.%y'))
with pd.ExcelWriter(BMpath) as writer:
	thirtyDay_490_final.to_excel(writer, sheet_name = '490 30 Day Inv', index = False)


print('30 Day files Saved')
