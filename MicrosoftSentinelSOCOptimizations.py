import json
import datetime
import toml
from quickchart  import QuickChart # sudo pip3 install quickchart.io
from docxtpl import DocxTemplate   #pip3 install docxtpl
from docx.shared import Mm
from docxtpl import InlineImage
import requests
import toml

Banner = """
  __  __ _                           __ _                        
 |  \/  (_) ___ _ __ ___  ___  ___  / _| |_                      
 | |\/| | |/ __| '__/ _ \/ __|/ _ \| |_| __|                     
 | |  | | | (__| | | (_) \__ \ (_) |  _| |_                      
 |_|__|_|_|\___|_| _\___/|___/\___/|_|  \__|   ___   ____        
 / ___|  ___ _ __ | |_(_)_ __   ___| | / ___| / _ \ / ___|       
 \___ \ / _ \ '_ \| __| | '_ \ / _ \ | \___ \| | | | |           
  ___) |  __/ | | | |_| | | | |  __/ |  ___) | |_| | |___        
 |____/ \___|_| |_|\__|_|_| |_|\___|_| |____/_\___/ \____|       
  / _ \ _ __ | |_(_)_ __ ___ (_)______ _| |_(_) ___  _ __  ___   
 | | | | '_ \| __| | '_ ` _ \| |_  / _` | __| |/ _ \| '_ \/ __|  
 | |_| | |_) | |_| | | | | | | |/ / (_| | |_| | (_) | | | \__ \  
  \___/| .__/ \__|_|_| |_| |_|_/___\__,_|\__|_|\___/|_| |_|___/  
       |_|                                                       

                Report Generator - 2025
"""

print(Banner)
# Load Configuration File
config = toml.load('Config/Config.toml')

Client_ID =  config['Client_ID']
Client_Secret =  config['Client_Secret']
EntraID_Tenant =  config['EntraID_Tenant']
Workspace =  config['Workspace']
WorkspaceID = config['WorkspaceID']
subscriptionID =  config['subscriptionID']
ResourceGroup = config['ResourceGroup']


# Get Microsoft Sentinel Access Token
def GetMicrosoftSentinelToken(Client_ID, Client_Secret, EntraID_Tenant):
    Url = "https://login.microsoftonline.com/"+EntraID_Tenant+"/oauth2/token"
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    payload='grant_type=client_credentials&client_id='+ Client_ID+'&resource=https%3A%2F%2Fmanagement.azure.com&client_secret='+Client_Secret
    response = requests.post(Url, headers=headers, data=payload).json()
    Access_Token = response["access_token"]
    print("[+] Access Token Received Successfully")
    print("[+] Connecting to Microsoft Sentinel ...")
    return Access_Token

# Print access token

Access_Token = GetMicrosoftSentinelToken(Client_ID, Client_Secret, EntraID_Tenant)

# Get Microsoft Sentinel SOC Optimization Recommendations
def GetMicrosoftSentinelRecommendations(Access_Token, Subscription, ResourceGroup, Workspace):
    Url = "https://management.azure.com/subscriptions/"+Subscription+"/resourceGroups/"+ResourceGroup+"/providers/Microsoft.OperationalInsights/workspaces/"+Workspace+"/providers/Microsoft.SecurityInsights/recommendations?api-version=2024-10-01-preview"
    Auth = 'Bearer '+Access_Token
    headers = {
      'Authorization': Auth ,
      'Content-Type': 'application/json',
    }
    response = requests.get(Url, headers=headers).json()
    print("[+] Microsoft Sentinel SOC Optimization Recommendations Received Successfully")
    return response


print("[+] Extracting Microsoft Sentinel SOC Optimization Recommendations ...")
Recommendations = GetMicrosoftSentinelRecommendations(Access_Token, subscriptionID,ResourceGroup, Workspace)


Optimizations = []
ActiveOptimizations = 0
InProgressOptimizations = 0
CompletedOptimizations = 0
DismissedOptimizations = 0
for optimization in Recommendations['value']:
    Title = optimization['properties']['title']
    State = optimization['properties']['state']
    Description = optimization['properties']['description']
    CreationTime = optimization['properties']['creationTimeUtc']
    if "Completed" in State:
        CompletedOptimizations += 1
    elif "Active" in State:
        ActiveOptimizations += 1
    elif "Progress" in State:
        InProgressOptimizations += 1
    elif "Dismissed" in State:
        DismissedOptimizations += 1
    # print("Title: ", Title)
    # print("Status: ", State)
    # print("Description: ", Description)
    # print("Creation Time: ", CreationTime)
    # print("-"*50)
    Optimizations.append({
        "Title": Title,
        "Status": State,
        "Description": Description,
        "CreationTime": CreationTime
    })

# Optimization Stats   
TotalOptimizations = len(Optimizations) 
print("[+] Total Optimizations: ", TotalOptimizations)
print("[+] Active Optimizations: ", ActiveOptimizations)
print("[+] In Progress Optimizations: ", InProgressOptimizations)
print("[+] Completed Optimizations: ", CompletedOptimizations)
print("[+] Dismissed Optimizations: ", DismissedOptimizations)

# Generate a pie chart for optimization stats
ChartResp = requests.post('https://quickchart.io/apex-charts/render', json={
    'width': 600,
    'height': 600,
    'config': "{ chart: { type: 'donut' }, series: "+str([ActiveOptimizations,InProgressOptimizations,CompletedOptimizations,DismissedOptimizations])+", labels: ['Active','In Progress','Completed','Dismissed'],colors:['#E52020', '#FF9D23', '#3A7D44','#27445D']  }",
})

with open('Resources/StatusChart.png', 'wb') as f:
    f.write(ChartResp.content)

# Current Month and Year
now = datetime.datetime.now()
Month = now.strftime("%B")
Year = now.strftime("%Y")
print("[+] Generating Microsoft Sentinel SOC Optimizations Report ...")
# Generate a Word Document
doc = DocxTemplate("Resources/ReportTemplate.docx")
context = {
    'DateYear': Month+" "+Year,
    'TotalOptimizations': TotalOptimizations,
    'ActiveOptimizations': ActiveOptimizations,
    'InProgressOptimizations': InProgressOptimizations,
    'CompletedOptimizations': CompletedOptimizations,
    'DismissedOptimizations': DismissedOptimizations,
    'Chart': InlineImage(doc, 'Resources/StatusChart.png', width=Mm(100)),
    'Optimizations': Optimizations
}
doc.render(context)
doc.save("MicrosoftSentinelSOCOptimizationsReport.docx")
print("[+] Microsoft Sentinel SOC Optimizations Report Generated Successfully")
print("[+] Report Saved as MicrosoftSentinelSOCOptimizationsReport.docx")