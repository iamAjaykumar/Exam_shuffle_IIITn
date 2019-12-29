import requests as re 
from openpyxl import load_workbook
from bs4 import BeautifulSoup
from openpyxl import Workbook
import pandas as pd 
import datetime 
date=datetime.date.today()
presentdate=date.strftime("%d-%m-%y")

try:

    writer = pd.ExcelWriter('shuffe.xlsx', engine='openpyxl') 
    wb  = writer.book
    columns=[]
    rows=[]
    batch=int(input("Enter the batch number: "))
    print("EX: N160218 \'16' is the batch number: ...")
    url="http://10.11.4.25/p2Shuffle/index.php"
    for dt in range(batch*10000,batch*10000+1300):
        

        #json fields to be sent to url 
        data={"ID": "N"+str(dt),
        "PWD":"", 
        "submit": "Submit"}
        #sending the data using session
        s=re.Session()
        output=s.post(url=url,data=data)
        #extracting the data field in html page
        soup = BeautifulSoup(output.content, 'html.parser')
        a=soup.find_all("b")
        li=[]
        for va in a:
            li.append((va.text.strip()))
        need=(li[4:])
        #print(need)
        for i in range(0,len(need)-1,2):
            rows.append((need[i+1]))

        
        #res_dct = {need[i]: need[i + 1] for i in range(0, len(need)-1, 2)} 
    for i in range(0,len(need)-1,2):
        columns.append(need[i])
    sid = [rows[i] for i in range(len(rows)) if i%7==0]
    sname=[rows[i] for i in range(len(rows)) if i%7==1]
    year=[rows[i] for i in range(len(rows)) if i%7==2]
    hall=[rows[i] for i in range(len(rows)) if i%7==4]
    desk=[rows[i] for i in range(len(rows)) if i%7==5]
    ip=[rows[i] for i in range(len(rows)) if i%7==6]


    #saving as Excel sheet

    df = pd.DataFrame({"Student ID": sid,
        'Student Name:': sname,
                    'Year and Branch': year,
                    
                    'Examination Hall': hall,
                    'Desk Position' :desk,
                    'IP Address:':ip })

    df.to_excel(writer, index=False)
    wb.save("N"+str(batch)+"_shuffle_%s.xlsx"%presentdate)
    print("Successfully created the Excel sheet :) ")

    #print(columns)
except:
    print("\n\n1.Make sure you are conncected to Rgukt Intranet.... \n2.Shuffle of :",presentdate,"Not yet posted :( ..)")