#removing every entry except "receipt"
from typing import Text
import xml.etree.ElementTree as ET
import csv

mytree = ET.parse('input.xml')
myroot = mytree.getroot()


for x in myroot[1][0][1].findall('TALLYMESSAGE'): #checking every 'TALLYMESSAGE' tag inside 'REQUESTDATA' 
    for y in x.findall('VOUCHER'): #checking 'VOUCHER' tag inside every 'TALLYMESSAGE' tag 
        vtype = y.find('VOUCHERTYPENAME').text  
        if(vtype!='Receipt'): #comparing the "VOUCHERTYPENAME" tag's contents 
            x.remove(y) #Removing of not of type 'Receeipt'

mytree.write('new.xml')

# from xml to csv
import pandas as pd
import xml.etree.ElementTree as Xet
  
cols = ["date", "Vch_No", "Particulars", "Vch_Type", "amount","Amount_Verified"]
rows = []
  
# Parsing the XML file
xmlparse = Xet.parse('new.xml')
root = xmlparse.getroot()

for j in root[1][0][1].findall("TALLYMESSAGE"):
    for i in j.findall("VOUCHER"):
        for k in i.findall("ALLLEDGERENTRIES.LIST"):
    
            date = i.find("DATE").text

            Vch_No = i.find("VOUCHERNUMBER").text

            Vch_Type = i.find("VOUCHERTYPENAME").text

            Particulars = i.find("PARTYLEDGERNAME").text

            amount = k.find("AMOUNT").text

            verif = i.find("ISVATDUTYPAID").text
    
            rows.append({"date": date,
                        "Vch_No": Vch_No,
                        "Particulars": Particulars,
                        "Vch_Type":Vch_Type,
                        "amount": amount,
                        "Amount_Verified":verif})
  
df = pd.DataFrame(rows, columns=cols)
  
# Writing dataframe to csv
df.to_csv('output.csv')


#for csv to xlsx
from openpyxl import Workbook
import csv


wb = Workbook()
ws = wb.active
with open('output.csv', 'r') as f:
    for row in csv.reader(f):
        ws.append(row)
wb.save('final.xlsx')