from simple_salesforce import Salesforce, SalesforceLogin
import requests
import pandas as pd
from io import StringIO



sf = Salesforce(username='nmondello@rjreliance.com.dev6',password='Monde100$', security_token='T8sz4CxJtqQfZaI38K8mks0e', domain = 'test')

query = """
SELECT Name, pse__Start_Date__c, pse__Estimated_Hours__c, pse__Project__r.pse__Opportunity__r.Sales_Channel__c FROM pse__Est_Vs_Actuals__c
"""
response = sf.query_all(query)
records = response['records']
temp = response['records']

# Create a DataFrame to store the retrieved data
df = pd.DataFrame(records)

# Select the desired columns
df = df[['Name', 'pse__Start_Date__c', 'pse__Estimated_Hours__c', 'pse__Project__r']]

for i in range(len(records)):
    try:
        df.loc[i, 'pse__Project__r'] =  str(temp[i]["pse__Project__r"]["pse__Opportunity__r"]["Sales_Channel__c"])
    except:
        df.loc[i, 'pse__Project__r'] = "No Opportunity"

# Divide the "Estimate_Hours__c" values by 40
df['pse__Estimated_Hours__c'] = df['pse__Estimated_Hours__c'] / 40.0

# Export the data to Excel
writer = pd.ExcelWriter('Salesforce2 Output.xlsx') #this will write the file to the same folder where this program is kept
df.to_excel(writer,index=False,header=True)
writer.close()