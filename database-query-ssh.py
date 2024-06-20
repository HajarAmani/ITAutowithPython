import pymysql
import paramiko
import pandas as pd
from paramiko import SSHClient
from sshtunnel import SSHTunnelForwarder
from os.path import expanduser
import time

sql_hostname = '000.00.000.00'
sql_username = 'me'
sql_password = 'me'
sql_main_database = 'main'
sql_port = 3306
ssh_host = '111.11.111.11'
ssh_user = 'me'
ssh_password = '****'
ssh_port = 22


df = pd.read_excel('C:\\Users\\me\\Locklist_Input.xlsx')
outdf = pd.read_excel('C:\\Users\\me\\Locklist_Out.xlsx')


for index, row in df.iterrows():
     lockID = row['key ID']
     with SSHTunnelForwarder(
        (ssh_host, ssh_port),
        ssh_username=ssh_user,
        ssh_password=ssh_password,
        remote_bind_address=(sql_hostname, sql_port)) as tunnel:
        conn = pymysql.connect(host='000.0.0.0', user=sql_username,
            passwd=sql_password, db=sql_main_database,
            port=tunnel.local_bind_port)
        query = "SELECT ai FROM db.a lo JOIN db.b lcd on a.ref_no = b.ref_no where json_extract(json_extract(json_extract(a.lock_information,'$.lockStockItem'),'$.inv_lockResult'),'$.inv1_lockId') like '%"+str(lockID)+"%'"
        ref_id = pd.read_sql_query(query, conn)
      
        if ref_id.empty == True:
             ref_id['ref_no'] = ['No Order']
        ref_id = ref_id.squeeze(axis=0) #series
        conn.close()
     row = row.append(ref_id)
     row = row.to_frame().transpose()
     outdf = pd.concat([outdf,row])
     print(str(index)+"/1648")
     

outdf.to_csv('C:\\Users\\me\\Locklist_Out_Final.csv',index=False,mode='a')