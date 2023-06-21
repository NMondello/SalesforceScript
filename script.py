from simple_salesforce import Salesforce, SalesforceLogin
import requests
import pandas as pd
from io import StringIO
import numpy as np



sf = Salesforce(username='nmondello@rjreliance.com.dev6',password='Monde100$', security_token='T8sz4CxJtqQfZaI38K8mks0e', domain = 'test')

#Query to get name, region, start date, estimated hours, sales channel
#__r signifies a relationship object
query = """
SELECT pse__Resource__r.Name, pse__Resource__r.pse__Region__r.Name, 
pse__Start_Date__c, pse__Estimated_Hours__c, pse__Actual_Hours__c, pse__Project__r.pse__Opportunity__r.Sales_Channel__c, pse__Project__r.pse__Is_Billable__c, Actual_Dollars__c, Estimated_Dollars__c FROM pse__Est_Vs_Actuals__c
WHERE pse__Start_Date__c > THIS_WEEK AND pse__Start_Date__c < NEXT_N_YEARS:2
"""
response = sf.query_all(query)
records = response['records']
temp = response['records']

# Create a DataFrame to store the retrieved data
df = pd.DataFrame(records)

# Select the desired columns and rename.
df = df[['pse__Resource__r', 'pse__Resource__r', 'pse__Start_Date__c', 'pse__Estimated_Hours__c', 'pse__Actual_Hours__c', 'pse__Project__r', 'pse__Project__r','Actual_Dollars__c', 'Estimated_Dollars__c']]
df.columns = ['Name', 'Region', 'Date', 'FTE-E', 'FTE-A', 'Type', 'Billable', 'Actual Dollars', 'Estimate Dollars']

#Loop for names, then sort alph
for i in range(len(temp)):
    try:
        df.loc[i, 'Name'] = temp[i]["pse__Resource__r"]['Name']
    except:
        df.loc[i, 'Name'] = 'N/A'

#Loop to get regions
for i in range(len(temp)):
    try:
        df.loc[i, 'Region'] = temp[i]["pse__Resource__r"]['pse__Region__r']["Name"]
    except:
        df.loc[i, 'Region'] = 'Unlisted'

#Loop to get billable
for i in range(len(temp)):
    try:
        df.loc[i, 'Billable'] = temp[i]["pse__Project__r"]['pse__Is_Billable__c']
    except:
        df.loc[i, 'Billable'] = 'Unlisted'

#Loop and identify types
for i in range(len(temp)):
    try:
        currentType = str(temp[i]["pse__Project__r"]["pse__Opportunity__r"]["Sales_Channel__c"])
        currentType2 = str(temp[i]["pse__Project__r"]["pse__Is_Billable__c"])
        if ('False' in currentType2):
            df.loc[i, 'Type'] = 'Non-Billable'
        elif ('Partner Referral - UKG Service' in currentType):
            df.loc[i, 'Type'] = 'UKG Services'
        elif ('Partner Referral - UKG' in currentType):
            df.loc[i, 'Type'] = 'UKG Direct Sales'
        elif ('SaaS Direct Sales' in currentType):
            df.loc[i, 'Type'] = 'SaaS Direct Sales'
        elif('Partner Referral - ADP' in currentType):
            df.loc[i, 'Type'] = 'ADP Services'
        else:
            df.loc[i, 'Type'] =  'All Other Direct Sales'
    except:
        currentType2 = str(temp[i]["pse__Project__r"]["pse__Is_Billable__c"])
        if ('False' in currentType2):
            df.loc[i, 'Type'] = 'Non-Billable'
        else:
            df.loc[i, 'Type'] = np.nan

#Get rid of jobs with no opportunity and that weren billable
df.dropna()

# Divide the "Estimate_Hours__c" values by 40
df['FTE-E'] = df['FTE-E'] / 40.0
df['FTE-A'] = df['FTE-A'] / 40.0

#Get rid of rows that have 0 hours worked.
df= df[df['FTE-E'] != 0]

#Sorts by week, and adds the FTE
df['Date'] = pd.to_datetime(df['Date']) - pd.to_timedelta(7, unit='d')
df = df.groupby(['Name', 'Region', pd.Grouper(key='Date', freq='W-MON'), pd.Grouper(key='Type')]).agg(
     FTE_Expected = ('FTE-E','sum'),
     FTE_Actual = ('FTE-A','sum'),
     Estimate_Dollars = ('Estimate Dollars', 'sum'),
     Actual_Dollars = ('Actual Dollars', 'sum'),
     ).reset_index().sort_values('Date')
df = df.sort_values(['Name', 'Date'])

#Get totals
df['Total FTE'] = df.apply(lambda x: x['FTE_Actual'] - x['FTE_Expected'], axis=1)
df['Total Dollars'] = df.apply(lambda x: x['Actual_Dollars'] - x['Estimate_Dollars'], axis=1)
# Export the data to Excel
writer = pd.ExcelWriter('Salesforce3 Output.xlsx') #this will write the file to the same folder where this program is kept
df.to_excel(writer,index=False,header=True)
writer.close()