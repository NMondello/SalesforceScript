from simple_salesforce import Salesforce, SalesforceLogin
import requests
import pandas as pd
from io import StringIO



sf = Salesforce(username='nmondello@rjreliance.com.dev6',password='Monde100$', security_token='T8sz4CxJtqQfZaI38K8mks0e', domain = 'test')

#Query to get name, region, start date, estimated hours, sales channel
#__r signifies a relationship object
query = """
SELECT pse__Resource__r.Name, pse__Project__r.pse__Opportunity__r.pse__Region__r.Name, pse__Start_Date__c, pse__Estimated_Hours__c, pse__Project__r.pse__Opportunity__r.Sales_Channel__c FROM pse__Est_Vs_Actuals__c
WHERE pse__Start_Date__c > THIS_WEEK AND pse__Start_Date__c < NEXT_N_YEARS:2
"""
response = sf.query_all(query)
records = response['records']
temp = response['records']

# Create a DataFrame to store the retrieved data
df = pd.DataFrame(records)

# Select the desired columns and rename.
df = df[['pse__Resource__r', 'pse__Project__r', 'pse__Start_Date__c', 'pse__Estimated_Hours__c', 'pse__Project__r']]
df.columns = ['Name', 'Region', 'Date', 'FTE', 'Type']


#Loop for names, then sort alph
for i in range(len(temp)):
    try:
        df.loc[i, 'Name'] = temp[i]["pse__Resource__r"]['Name']
    except:
        df.loc[i, 'Name'] = 'N/A'

#Loop to get regions
for i in range(len(temp)):
    try:
        df.loc[i, 'Region'] = temp[i]["pse__Project__r"]["pse__Opportunity__r"]['pse__Region__r']["Name"]
    except:
        df.loc[i, 'Region'] = 'Unlisted'

#Loop and identify types
for i in range(len(temp)):
    try:
        currentType = str(temp[i]["pse__Project__r"]["pse__Opportunity__r"]["Sales_Channel__c"])
        if ('Partner Referral - ADP' in currentType):
            df.loc[i, 'Type'] = 'ADP Services'
        elif ('Partner Referral - UKG Service' in currentType):
            df.loc[i, 'Type'] = 'UKG Services'
        elif ('Partner Referral - UKG' in currentType):
            df.loc[i, 'Type'] = 'UKG Direct Sales'
        elif ('SaaS Direct Sales' in currentType):
            df.loc[i, 'Type'] = 'SaaS Direct Sales'
        else:
            df.loc[i, 'Type'] =  'All Other Direct Sales'
    except:
        df.loc[i, 'Type'] = "No Opportunity"

# Divide the "Estimate_Hours__c" values by 40
df['FTE'] = df['FTE'] / 40.0

#Sorts by week, and adds the FTE
df['Date'] = pd.to_datetime(df['Date']) - pd.to_timedelta(7, unit='d')
df = df.groupby(['Name', 'Region', pd.Grouper(key='Date', freq='W-MON'), pd.Grouper(key='Type')])['FTE'].sum().reset_index().sort_values('Date')
df = df.sort_values(['Name', 'Date'])
# Export the data to Excel
writer = pd.ExcelWriter('Salesforce3 Output.xlsx') #this will write the file to the same folder where this program is kept
df.to_excel(writer,index=False,header=True)
writer.close()