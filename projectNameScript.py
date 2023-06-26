from simple_salesforce import Salesforce, SalesforceLogin
import requests
import pandas as pd
from io import StringIO



sf = Salesforce(username='USER',password='PASS', security_token='TOKEN', domain = 'test')

#Query to get name, region, start date, estimated hours, sales channel
#__r signifies a relationship object
query = """
SELECT pse__Project__r.Name, pse__Project__r.pse__Opportunity__r.Name FROM pse__Est_Vs_Actuals__c

"""
response = sf.query_all(query)
records = response['records']
temp = response['records']
df = pd.DataFrame(records)

df = df[['pse__Project__r', 'pse__Project__r']]
df.columns = ['Project Name', 'Status']

#Loop for names, then sort alph
for i in range(len(temp)):
    try:
        df.loc[i, 'Project Name'] = temp[i]["pse__Project__r"]['Name']
    except:
        df.loc[i, 'Project Name'] = 'N/A'

for i in range(len(temp)):
    try:
        currentType = str(temp[i]["pse__Project__r"]["pse__Opportunity__r"])
        if ('None' in currentType):
            df.loc[i, 'Status'] = 'PROJECT WITH NO OPP'
        else:
            df.loc[i, 'Status'] = str(temp[i]["pse__Project__r"]["pse__Opportunity__r"]['Name'])
    except:
        df.loc[i, 'Status'] = "No Opportunity FAILED"

writer = pd.ExcelWriter('Project Output.xlsx') #this will write the file to the same folder where this program is kept
df.to_excel(writer,index=False,header=True)
writer.close()
