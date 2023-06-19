from simple_salesforce import Salesforce, SalesforceLogin
import requests
import pandas as pd
from io import StringIO
# only want going forward in the future 2 years
sf = Salesforce(username='nmondello@rjreliance.com.dev6',password='Monde100$', security_token='T8sz4CxJtqQfZaI38K8mks0e', domain = 'test')
query = "SELECT pse__Project__r.pse__Opportunity__r.Sales_Channel__c FROM pse__Est_Vs_Actuals__c" #join to project table and join back to opportunity table, then pull forward sales channel
response = sf.query(query)
records = response["records"]

for i in range(len(records)):
    try:
        print(records[i]["pse__Project__r"]["pse__Opportunity__r"]["Sales_Channel__c"])
    except:
        print("No Opportunity")

#for record in records:
    #name = record['Name']
    #estimate_hours = record['pse__Estimated_Hours__c']
    # Perform operations on the retrieved values
    #FTI = estimate_hours/40.0
    # Example: print the name and estimate hours
    #print(f"Item Name: {name}, FTI: {FTI}")

# Create a DataFrame to store the retrieved data
# df = pd.DataFrame(records)
#
# # Select the desired columns
# df = df[['Name', 'pse__Project__c']]
#
# # Divide the "Estimate_Hours__c" values by 40
# #df['pse__Estimated_Hours__c'] = df['pse__Estimated_Hours__c'] / 40.0
#
# # Export the data to Excel
# writer = pd.ExcelWriter('Salesforce Output.xlsx') #this will write the file to the same folder where this program is kept
# df.to_excel(writer,index=False,header=True)
