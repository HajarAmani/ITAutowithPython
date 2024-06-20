import requests
import pandas as pd
from datetime import datetime
import xml.etree.ElementTree as ET
from requests.packages.urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)



df = pd.read_excel('C:\\Users\\XXX.xlsx', sheet_name="Sheet1")
orgList = pd.read_excel('C:\\Users\\YYY.xlsx')

url = "https://00.00.00.0:00000/apiaccess/openAPI/XXX/XXX"

#generate token in postman
headers = {
     'Authorization': "Bearer XXX",
     'Content-Type': "application/xml"
     }

#modem unlock
for index, row in df.iterrows():
     now = datetime.now()
     sysdate = now.strftime("%Y%m%d%H%M%S")
     lockId = row['OPER_NOTE_ID']
     channel = row['EMPLOYEE_CODE']
     warehouseCode = row['IM_RWH_CODE']
     for index2, row2 in orgList.iterrows():
          if row2['IM_RWH_CODE'] == warehouseCode:
               orgId = row2['ORG_ID']
               break
     ouId = orgId
     payload = f'''<?xml version=\"1.0\"?>\r\n<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\">\r\n\t<soapenv:Header/>\r\n\t
     <soapenv:Body>\r\n\t\t<inv:UnLockReqMsg>\r\n\t\t\t<requestHeader>\r\n\t\t\t\t
     <crm:version>1.00</crm:version>\r\n\t\t\t\t<crm:messageSeq>{sysdate}</crm:messageSeq><crm:var1>XXXX</crm:var1>\r\n\t\t\t</requestHeader>
     \r\n\t</soapenv:Body>\r\n
     </soapenv:Envelope>'''
     response = requests.post(url, headers=headers, data=payload, verify=False)
     if response.status_code == 200:
          responseXml = ET.fromstring(response.text)
          successMsg = responseXml[1][0][0][2].text
          if successMsg != "success":
               print(str(index+1)+' no success msg '+str(lockId))
               f = open("C:\\Users\\me\\OneDrive\resultsuccess.txt", "a")
               f.write(str(row.to_dict()) + '\n')
               f.close()
     else:
          f = open("C:\\Users\\me\\OneDrive\resultfail.txt", "a")
          f.write(str(row.to_dict()) + '\n')
          f.close()
     print(str(index+1)+'/' + str(df.shape[0]))