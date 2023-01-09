# -*- coding: utf-8 -*-
"""
HealthEdge Benefit Insurance Template_POC
"""

##Main code
import pandas as pd
pd.set_option('display.max_columns', None)
import numpy as np
from PyPDF2 import PdfFileReader


# creating a pdf file object
pdfObject = open('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/2021-SBC-BCBSAZ-HSA.pdf', 'rb')

# creating a pdf reader object
pdfReader = PdfFileReader(pdfObject)

pdfpagelist = PdfFileReader(pdfObject).getPage(9).extractText().split('\n')
pdfpagelist


plan_plan = pd.read_excel('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx', sheet_name='Plan Name')
plan_plan['Detail']=plan_plan['Detail'].astype(str)

plan_plan.at[0, 'Detail'] = pdfpagelist[37][36:42]

pdfpagelist = PdfFileReader(pdfObject).getPage(0).extractText().split('\n')

plan_plan.at[1, 'Detail'] = pdfpagelist[2][188:205]
plan_plan.at[2, 'Detail'] = pdfpagelist[2][202:205]
plan_plan.at[3, 'Detail'] = 'N/A'


plan_detail = pd.read_excel('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx', sheet_name='Plan Details')
plan_detail=plan_detail.astype(str)
for i in range(len(plan_detail)):
    plan_detail.at[i, 'IN Copay'] = 'N/A' 
    plan_detail.at[i, 'IN Coins'] = 'N/A' 
    plan_detail.at[i, 'IN Coins After Deductible Flag'] = 'N/A' 
    plan_detail.at[i, 'OON Copay'] = 'N/A' 
    plan_detail.at[i, 'OON Coins'] = 'N/A' 
    plan_detail.at[i, 'OON Coins After Deductible Flag'] = 'N/A' 
    plan_detail.at[i, 'IN DED'] = 'N/A' 
    plan_detail.at[i, 'IN Max'] = 'N/A' 
    plan_detail.at[i, 'OON DED'] = 'N/A' 
    plan_detail.at[i, 'OON Max'] = 'N/A' 
    plan_detail.at[i, 'COVERED IN/OON/BOTH'] = 'N/A' 

    
for i in (2,9):
    plan_detail.at[i, 'IN Copay'] = '' 
    plan_detail.at[i, 'IN Coins'] = '' 
    plan_detail.at[i, 'IN Coins After Deductible Flag'] = '' 
    plan_detail.at[i, 'OON Copay'] = '' 
    plan_detail.at[i, 'OON Coins'] = '' 
    plan_detail.at[i, 'OON Coins After Deductible Flag'] = '' 
    plan_detail.at[i, 'IN DED'] = '' 
    plan_detail.at[i, 'IN Max'] = '' 
    plan_detail.at[i, 'OON DED'] = '' 
    plan_detail.at[i, 'OON Max'] = '' 
    plan_detail.at[i, 'COVERED IN/OON/BOTH'] = ''



pdfpagelist = PdfFileReader(pdfObject).getPage(0).extractText().split('\n')


plan_detail.at[0, 'IN DED'] = pdfpagelist[14][24:30]
plan_detail.at[1, 'IN DED'] = pdfpagelist[18][23:29]
plan_detail.at[0, 'IN Max'] = pdfpagelist[36][24:30]
plan_detail.at[1, 'IN Max'] = pdfpagelist[39][24:30]
plan_detail.at[0, 'OON DED'] = pdfpagelist[15][28:34]
plan_detail.at[1, 'OON DED'] = pdfpagelist[19][28:34]
plan_detail.at[0, 'OON Max'] = pdfpagelist[37][28:34]
plan_detail.at[1, 'OON Max'] = pdfpagelist[40][28:35]


pdfpagelist = PdfFileReader(pdfObject).getPage(1).extractText().split('\n')

pdfpagelist[33][14:23]

plan_detail.at[3, 'IN Coins'] = pdfpagelist[23][22:25]
plan_detail.at[4, 'IN Coins'] = pdfpagelist[23][22:25]
plan_detail.at[5, 'IN Coins'] = pdfpagelist[23][22:25]
plan_detail.at[6, 'IN Coins'] = pdfpagelist[23][22:25]
plan_detail.at[7, 'IN Coins'] = pdfpagelist[23][22:25]
plan_detail.at[8, 'IN Coins'] = pdfpagelist[33][14:23]


pdfpagelist = PdfFileReader(pdfObject).getPage(3).extractText().split('\n')


plan_detail.at[10, 'IN Coins'] = pdfpagelist[7][7:10]
plan_detail.at[11, 'IN Coins'] = pdfpagelist[7][7:10]

plan_detail.at[3, 'IN Coins After Deductible Flag'] = 'Yes'
plan_detail.at[4, 'IN Coins After Deductible Flag'] = 'Yes'
plan_detail.at[5, 'IN Coins After Deductible Flag'] = 'Yes'
plan_detail.at[6, 'IN Coins After Deductible Flag'] = 'Yes'
plan_detail.at[7, 'IN Coins After Deductible Flag'] = 'Yes'
plan_detail.at[8, 'IN Coins After Deductible Flag'] = 'No'
plan_detail.at[10, 'IN Coins After Deductible Flag'] = 'Yes'
plan_detail.at[11, 'IN Coins After Deductible Flag'] = 'Yes'

plan_detail.at[3, 'OON Coins'] = pdfpagelist[7][24:27]
plan_detail.at[4, 'OON Coins'] = pdfpagelist[7][24:27]
plan_detail.at[5, 'OON Coins'] = pdfpagelist[7][24:27]
plan_detail.at[6, 'OON Coins'] = pdfpagelist[7][24:27]
plan_detail.at[7, 'OON Coins'] = pdfpagelist[7][24:27]
plan_detail.at[8, 'OON Coins'] = pdfpagelist[7][24:27]
plan_detail.at[10, 'OON Coins'] = pdfpagelist[7][24:27]
plan_detail.at[11, 'OON Coins'] = pdfpagelist[7][7:10]

plan_detail.at[3, 'OON Coins After Deductible Flag'] = 'Yes'
plan_detail.at[4, 'OON Coins After Deductible Flag'] = 'Yes'
plan_detail.at[5, 'OON Coins After Deductible Flag'] = 'Yes'
plan_detail.at[6, 'OON Coins After Deductible Flag'] = 'Yes'
plan_detail.at[7, 'OON Coins After Deductible Flag'] = 'Yes'
plan_detail.at[8, 'OON Coins After Deductible Flag'] = 'Yes'
plan_detail.at[10, 'OON Coins After Deductible Flag'] = 'Yes'
plan_detail.at[11, 'OON Coins After Deductible Flag'] = 'Yes'



plan_detail.at[3, 'COVERED IN/OON/BOTH'] = np.where((plan_detail.at[3,'IN Copay']!= 'N/A' or plan_detail.at[3,'IN Coins']!= 'N/A') 
                                                    & (plan_detail.at[3,'OON Copay']!= 'N/A' or plan_detail.at[3,'OON Coins']!= 'N/A')
                                                    & (plan_detail.at[3,'OON Copay']!= 'No benefit' or plan_detail.at[3,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
                                                    
plan_detail.at[4, 'COVERED IN/OON/BOTH'] = np.where((plan_detail.at[4,'IN Copay']!= 'N/A' or plan_detail.at[4,'IN Coins']!= 'N/A') 
                                                    & (plan_detail.at[4,'OON Copay']!= 'N/A' or plan_detail.at[4,'OON Coins']!= 'N/A')
                                                    & (plan_detail.at[4,'OON Copay']!= 'No benefit' and plan_detail.at[4,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

plan_detail.at[5, 'COVERED IN/OON/BOTH'] = np.where((plan_detail.at[5,'IN Copay']!= 'N/A' or plan_detail.at[5,'IN Coins']!= 'N/A') 
                                                    & (plan_detail.at[5,'OON Copay']!= 'N/A' or plan_detail.at[5,'OON Coins']!= 'N/A')
                                                    & (plan_detail.at[5,'OON Copay']!= 'No benefit' and plan_detail.at[5,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

plan_detail.at[6, 'COVERED IN/OON/BOTH'] = np.where((plan_detail.at[6,'IN Copay']!= 'N/A' or plan_detail.at[6,'IN Coins']!= 'N/A') 
                                                    & (plan_detail.at[6,'OON Copay']!= 'N/A' or plan_detail.at[6,'OON Coins']!= 'N/A')
                                                    & (plan_detail.at[6,'OON Copay']!= 'No benefit' or plan_detail.at[6,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
                                                    
plan_detail.at[7, 'COVERED IN/OON/BOTH'] = np.where((plan_detail.at[7,'IN Copay']!= 'N/A' or plan_detail.at[7,'IN Coins']!= 'N/A') 
                                                    & (plan_detail.at[7,'OON Copay']!= 'N/A' or plan_detail.at[7,'OON Coins']!= 'N/A')
                                                    & (plan_detail.at[7,'OON Copay']!= 'No benefit' and plan_detail.at[7,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

plan_detail.at[8, 'COVERED IN/OON/BOTH'] = np.where((plan_detail.at[8,'IN Copay']!= 'N/A' or plan_detail.at[8,'IN Coins']!= 'N/A') 
                                                    & (plan_detail.at[8,'OON Copay']!= 'N/A' or plan_detail.at[8,'OON Coins']!= 'N/A')
                                                    & (plan_detail.at[8,'OON Copay']!= 'No benefit' and plan_detail.at[8,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

plan_detail.at[10, 'COVERED IN/OON/BOTH'] = np.where((plan_detail.at[10,'IN Copay']!= 'N/A' or plan_detail.at[10,'IN Coins']!= 'N/A') 
                                                    & (plan_detail.at[10,'OON Copay']!= 'N/A' or plan_detail.at[10,'OON Coins']!= 'N/A')
                                                    & (plan_detail.at[10,'OON Copay']!= 'No benefit' and plan_detail.at[10,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
plan_detail.at[11, 'COVERED IN/OON/BOTH'] = np.where((plan_detail.at[11,'IN Copay']!= 'N/A' or plan_detail.at[11,'IN Coins']!= 'N/A') 
                                                    & (plan_detail.at[11,'OON Copay']!= 'N/A' or plan_detail.at[11,'OON Coins']!= 'N/A')
                                                    & (plan_detail.at[11,'OON Copay']!= 'No benefit' and plan_detail.at[11,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')


path = 'C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/BCBS_AZ_HSA_Qualified_PPO.xlsx'

with pd.ExcelWriter(path) as writer:
    plan_plan.to_excel(writer, sheet_name='BCBSAZ',index = False)
    plan_detail.to_excel(writer,sheet_name = 'BCBSAZ',startrow = 5,index = False)
