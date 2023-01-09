# -*- coding: utf-8 -*-
"""
HealthEdge Benefit Insurance Template_POC
"""

##Main code
import pandas as pd
import numpy as np
import tabula
from PyPDF2 import PdfFileReader


data = tabula.read_pdf("C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/UHC-ca-detailed-benefit-grid-january-2021.pdf", pages = 9)
list = data[0].columns
arr = np.array(data)
df = pd.DataFrame(arr.reshape(-1,5))


uhc1_plan = pd.read_excel('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx', sheet_name='Plan Name')
uhc1_plan['Detail']=uhc1_plan['Detail'].astype(str)
uhc1_plan


uhc1_plan.at[0, 'Detail'] = 'UHC'
uhc1_plan.at[1, 'Detail'] = list[0][:16]
uhc1_plan.at[2, 'Detail'] = df.iloc[1][0]
uhc1_plan.at[3, 'Detail'] = df.iloc[0][1]

uhc1_detail = pd.read_excel('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx', sheet_name='Plan Details')
uhc1_detail=uhc1_detail.astype(str)
for i in range(len(uhc1_detail)):
    uhc1_detail.at[i, 'IN Copay'] = 'N/A' 
    uhc1_detail.at[i, 'IN Coins'] = 'N/A' 
    uhc1_detail.at[i, 'IN Coins After Deductible Flag'] = 'N/A' 
    uhc1_detail.at[i, 'OON Copay'] = 'N/A' 
    uhc1_detail.at[i, 'OON Coins'] = 'N/A' 
    uhc1_detail.at[i, 'OON Coins After Deductible Flag'] = 'N/A' 
    uhc1_detail.at[i, 'IN DED'] = 'N/A' 
    uhc1_detail.at[i, 'IN Max'] = 'N/A' 
    uhc1_detail.at[i, 'OON DED'] = 'N/A' 
    uhc1_detail.at[i, 'OON Max'] = 'N/A' 
    uhc1_detail.at[i, 'COVERED IN/OON/BOTH'] = 'N/A' 

    
for i in (2,9):
    uhc1_detail.at[i, 'IN Copay'] = '' 
    uhc1_detail.at[i, 'IN Coins'] = '' 
    uhc1_detail.at[i, 'IN Coins After Deductible Flag'] = '' 
    uhc1_detail.at[i, 'OON Copay'] = '' 
    uhc1_detail.at[i, 'OON Coins'] = '' 
    uhc1_detail.at[i, 'OON Coins After Deductible Flag'] = '' 
    uhc1_detail.at[i, 'IN DED'] = '' 
    uhc1_detail.at[i, 'IN Max'] = '' 
    uhc1_detail.at[i, 'OON DED'] = '' 
    uhc1_detail.at[i, 'OON Max'] = '' 
    uhc1_detail.at[i, 'COVERED IN/OON/BOTH'] = ''  
    
    
    
uhc1_detail.at[0, 'IN Max'] = df.iloc[4][0][:6]
uhc1_detail.at[1, 'IN Max'] = df.iloc[4][0][7:]
uhc1_detail.at[0, 'OON DED'] = df.iloc[3][1][:6]
uhc1_detail.at[1, 'OON DED'] = df.iloc[3][1][7:]
uhc1_detail.at[0, 'OON Max'] = df.iloc[4][1][:6]
uhc1_detail.at[1, 'OON Max'] = df.iloc[4][1][8:]


uhc1_detail.at[3, 'IN Copay'] = df.iloc[6][0]
uhc1_detail.at[4, 'IN Copay'] = df.iloc[7][0]
uhc1_detail.at[5, 'IN Copay'] = df.iloc[8][0]
uhc1_detail.at[6, 'IN Copay'] = df.iloc[9][0]
uhc1_detail.at[7, 'IN Copay'] = df.iloc[10][0]
uhc1_detail.at[8, 'IN Copay'] = df.iloc[11][0]
uhc1_detail.at[11, 'IN Copay'] = df.iloc[17][0]

uhc1_detail.at[10, 'IN Coins'] = df.iloc[13][0][:3]


uhc1_detail.at[10, 'IN Coins After Deductible Flag'] = 'No'


uhc1_detail.at[11, 'OON Copay'] = df.iloc[17][0]

uhc1_detail.at[3, 'OON Coins'] = df.iloc[6][3][:3]
uhc1_detail.at[4, 'OON Coins'] = df.iloc[7][3][:3]
uhc1_detail.at[5, 'OON Coins'] = df.iloc[8][3]
uhc1_detail.at[6, 'OON Coins'] = df.iloc[9][3][:3]
uhc1_detail.at[7, 'OON Coins'] = df.iloc[10][3][:3]
uhc1_detail.at[8, 'OON Coins'] = df.iloc[11][3]

uhc1_detail.at[10, 'OON Coins'] = df.iloc[13][3][:3]

uhc1_detail.at[3, 'OON Coins After Deductible Flag'] = 'Yes'
uhc1_detail.at[4, 'OON Coins After Deductible Flag'] = 'Yes'
uhc1_detail.at[6, 'OON Coins After Deductible Flag'] = 'Yes'
uhc1_detail.at[7, 'OON Coins After Deductible Flag'] = 'Yes'
uhc1_detail.at[10, 'OON Coins After Deductible Flag'] = 'Yes'


uhc1_detail.at[3, 'COVERED IN/OON/BOTH'] = np.where((uhc1_detail.at[3,'IN Copay']!= 'N/A' or uhc1_detail.at[3,'IN Coins']!= 'N/A') 
                                                    & (uhc1_detail.at[3,'OON Copay']!= 'N/A' or uhc1_detail.at[3,'OON Coins']!= 'N/A')
                                                    & (uhc1_detail.at[3,'OON Copay']!= 'No benefit' or uhc1_detail.at[3,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
                                                    
uhc1_detail.at[4, 'COVERED IN/OON/BOTH'] = np.where((uhc1_detail.at[4,'IN Copay']!= 'N/A' or uhc1_detail.at[4,'IN Coins']!= 'N/A') 
                                                    & (uhc1_detail.at[4,'OON Copay']!= 'N/A' or uhc1_detail.at[4,'OON Coins']!= 'N/A')
                                                    & (uhc1_detail.at[4,'OON Copay']!= 'No benefit' and uhc1_detail.at[4,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

uhc1_detail.at[5, 'COVERED IN/OON/BOTH'] = np.where((uhc1_detail.at[5,'IN Copay']!= 'N/A' or uhc1_detail.at[5,'IN Coins']!= 'N/A') 
                                                    & (uhc1_detail.at[5,'OON Copay']!= 'N/A' or uhc1_detail.at[5,'OON Coins']!= 'N/A')
                                                    & (uhc1_detail.at[5,'OON Copay']!= 'No benefit' and uhc1_detail.at[5,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

uhc1_detail.at[6, 'COVERED IN/OON/BOTH'] = np.where((uhc1_detail.at[6,'IN Copay']!= 'N/A' or uhc1_detail.at[6,'IN Coins']!= 'N/A') 
                                                    & (uhc1_detail.at[6,'OON Copay']!= 'N/A' or uhc1_detail.at[6,'OON Coins']!= 'N/A')
                                                    & (uhc1_detail.at[6,'OON Copay']!= 'No benefit' or uhc1_detail.at[6,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
                                                    
uhc1_detail.at[7, 'COVERED IN/OON/BOTH'] = np.where((uhc1_detail.at[7,'IN Copay']!= 'N/A' or uhc1_detail.at[7,'IN Coins']!= 'N/A') 
                                                    & (uhc1_detail.at[7,'OON Copay']!= 'N/A' or uhc1_detail.at[7,'OON Coins']!= 'N/A')
                                                    & (uhc1_detail.at[7,'OON Copay']!= 'No benefit' and uhc1_detail.at[7,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

uhc1_detail.at[8, 'COVERED IN/OON/BOTH'] = np.where((uhc1_detail.at[8,'IN Copay']!= 'N/A' or uhc1_detail.at[8,'IN Coins']!= 'N/A') 
                                                    & (uhc1_detail.at[8,'OON Copay']!= 'N/A' or uhc1_detail.at[8,'OON Coins']!= 'N/A')
                                                    & (uhc1_detail.at[8,'OON Copay']!= 'No benefit' and uhc1_detail.at[8,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

uhc1_detail.at[10, 'COVERED IN/OON/BOTH'] = np.where((uhc1_detail.at[10,'IN Copay']!= 'N/A' or uhc1_detail.at[10,'IN Coins']!= 'N/A') 
                                                    & (uhc1_detail.at[10,'OON Copay']!= 'N/A' or uhc1_detail.at[10,'OON Coins']!= 'N/A')
                                                    & (uhc1_detail.at[10,'OON Copay']!= 'No benefit' and uhc1_detail.at[10,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
uhc1_detail.at[11, 'COVERED IN/OON/BOTH'] = np.where((uhc1_detail.at[11,'IN Copay']!= 'N/A' or uhc1_detail.at[11,'IN Coins']!= 'N/A') 
                                                    & (uhc1_detail.at[11,'OON Copay']!= 'N/A' or uhc1_detail.at[11,'OON Coins']!= 'N/A')
                                                    & (uhc1_detail.at[11,'OON Copay']!= 'No benefit' and uhc1_detail.at[11,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')


path = 'C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/UHC_Example_1_completed.xlsx'

with pd.ExcelWriter(path) as writer:
    uhc1_plan.to_excel(writer, sheet_name='Example1',index = False)
    uhc1_detail.to_excel(writer,sheet_name = 'Example1',startrow = 5,index = False)
    
    
    
    
data = tabula.read_pdf("C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/UHC-ca-detailed-benefit-grid-january-2021.pdf", pages = 9)
list = data[0].columns
arr = np.array(data)
df = pd.DataFrame(arr.reshape(-1,5))
df   


uhc2_plan = pd.read_excel('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx', sheet_name='Plan Name')
uhc2_plan['Detail']=uhc2_plan['Detail'].astype(str)
uhc2_plan


uhc2_plan.at[0, 'Detail'] = 'UHC'
uhc2_plan.at[1, 'Detail'] = list[0][:16]
uhc2_plan.at[2, 'Detail'] = df.iloc[1][0]
uhc2_plan.at[3, 'Detail'] = df.iloc[0][2]



uhc2_detail = pd.read_excel('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx', sheet_name='Plan Details')
uhc2_detail=uhc2_detail.astype(str)
for i in range(len(uhc2_detail)):
    uhc2_detail.at[i, 'IN Copay'] = 'N/A' 
    uhc2_detail.at[i, 'IN Coins'] = 'N/A' 
    uhc2_detail.at[i, 'IN Coins After Deductible Flag'] = 'N/A' 
    uhc2_detail.at[i, 'OON Copay'] = 'N/A' 
    uhc2_detail.at[i, 'OON Coins'] = 'N/A' 
    uhc2_detail.at[i, 'OON Coins After Deductible Flag'] = 'N/A' 
    uhc2_detail.at[i, 'IN DED'] = 'N/A' 
    uhc2_detail.at[i, 'IN Max'] = 'N/A' 
    uhc2_detail.at[i, 'OON DED'] = 'N/A' 
    uhc2_detail.at[i, 'OON Max'] = 'N/A' 
    uhc2_detail.at[i, 'COVERED IN/OON/BOTH'] = 'N/A' 

    
for i in (2,9):
    uhc2_detail.at[i, 'IN Copay'] = '' 
    uhc2_detail.at[i, 'IN Coins'] = '' 
    uhc2_detail.at[i, 'IN Coins After Deductible Flag'] = '' 
    uhc2_detail.at[i, 'OON Copay'] = '' 
    uhc2_detail.at[i, 'OON Coins'] = '' 
    uhc2_detail.at[i, 'OON Coins After Deductible Flag'] = '' 
    uhc2_detail.at[i, 'IN DED'] = '' 
    uhc2_detail.at[i, 'IN Max'] = '' 
    uhc2_detail.at[i, 'OON DED'] = '' 
    uhc2_detail.at[i, 'OON Max'] = '' 
    uhc2_detail.at[i, 'COVERED IN/OON/BOTH'] = ''    


uhc2_detail.at[0, 'IN DED'] = df.iloc[3][2][:4]
uhc2_detail.at[1, 'IN DED'] = df.iloc[3][2][5:]
uhc2_detail.at[0, 'IN Max'] = df.iloc[4][2][:6]
uhc2_detail.at[1, 'IN Max'] = df.iloc[4][2][7:]
uhc2_detail.at[0, 'OON DED'] = df.iloc[3][3][:6]
uhc2_detail.at[1, 'OON DED'] = df.iloc[3][3][7:]
uhc2_detail.at[0, 'OON Max'] = df.iloc[4][3][:7]
uhc2_detail.at[1, 'OON Max'] = df.iloc[4][3][8:]


uhc2_detail.at[3, 'IN Copay'] = df.iloc[6][2]
uhc2_detail.at[4, 'IN Copay'] = df.iloc[7][2]
uhc2_detail.at[5, 'IN Copay'] = df.iloc[8][2]
uhc2_detail.at[6, 'IN Copay'] = df.iloc[9][2]
uhc2_detail.at[7, 'IN Copay'] = df.iloc[10][2]
uhc2_detail.at[8, 'IN Copay'] = df.iloc[11][2]

uhc2_detail.at[10, 'IN Coins'] = df.iloc[13][2][:3]
uhc2_detail.at[11, 'IN Coins'] = df.iloc[17][2][:3]

uhc2_detail.at[10, 'IN Coins After Deductible Flag'] = 'Yes'
uhc2_detail.at[11, 'IN Coins After Deductible Flag'] = 'Yes'

uhc2_detail.at[3, 'OON Coins'] = df.iloc[6][3][:3]
uhc2_detail.at[4, 'OON Coins'] = df.iloc[7][3][:3]
uhc2_detail.at[5, 'OON Coins'] = df.iloc[8][3]
uhc2_detail.at[6, 'OON Coins'] = df.iloc[9][3][:3]
uhc2_detail.at[7, 'OON Coins'] = df.iloc[10][3][:3]
uhc2_detail.at[8, 'OON Coins'] = df.iloc[11][3]

uhc2_detail.at[10, 'OON Coins'] = df.iloc[13][3][:3]
uhc2_detail.at[11, 'OON Coins'] = df.iloc[17][2][:3]

uhc2_detail.at[3, 'OON Coins After Deductible Flag'] = 'Yes'
uhc2_detail.at[4, 'OON Coins After Deductible Flag'] = 'Yes'
uhc2_detail.at[6, 'OON Coins After Deductible Flag'] = 'Yes'
uhc2_detail.at[7, 'OON Coins After Deductible Flag'] = 'Yes'
uhc2_detail.at[10, 'OON Coins After Deductible Flag'] = 'Yes'
uhc2_detail.at[11, 'OON Coins After Deductible Flag'] = 'Yes'

uhc2_detail.at[3, 'COVERED IN/OON/BOTH'] = np.where((uhc2_detail.at[3,'IN Copay']!= 'N/A' or uhc2_detail.at[3,'IN Coins']!= 'N/A') 
                                                    & (uhc2_detail.at[3,'OON Copay']!= 'N/A' or uhc2_detail.at[3,'OON Coins']!= 'N/A')
                                                    & (uhc2_detail.at[3,'OON Copay']!= 'No benefit' or uhc2_detail.at[3,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
                                                    
uhc2_detail.at[4, 'COVERED IN/OON/BOTH'] = np.where((uhc2_detail.at[4,'IN Copay']!= 'N/A' or uhc2_detail.at[4,'IN Coins']!= 'N/A') 
                                                    & (uhc2_detail.at[4,'OON Copay']!= 'N/A' or uhc2_detail.at[4,'OON Coins']!= 'N/A')
                                                    & (uhc2_detail.at[4,'OON Copay']!= 'No benefit' and uhc2_detail.at[4,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

uhc2_detail.at[5, 'COVERED IN/OON/BOTH'] = np.where((uhc2_detail.at[5,'IN Copay']!= 'N/A' or uhc2_detail.at[5,'IN Coins']!= 'N/A') 
                                                    & (uhc2_detail.at[5,'OON Copay']!= 'N/A' or uhc2_detail.at[5,'OON Coins']!= 'N/A')
                                                    & (uhc2_detail.at[5,'OON Copay']!= 'No benefit' and uhc2_detail.at[5,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

uhc2_detail.at[6, 'COVERED IN/OON/BOTH'] = np.where((uhc2_detail.at[6,'IN Copay']!= 'N/A' or uhc2_detail.at[6,'IN Coins']!= 'N/A') 
                                                    & (uhc2_detail.at[6,'OON Copay']!= 'N/A' or uhc2_detail.at[6,'OON Coins']!= 'N/A')
                                                    & (uhc2_detail.at[6,'OON Copay']!= 'No benefit' or uhc2_detail.at[6,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
                                                    
uhc2_detail.at[7, 'COVERED IN/OON/BOTH'] = np.where((uhc2_detail.at[7,'IN Copay']!= 'N/A' or uhc2_detail.at[7,'IN Coins']!= 'N/A') 
                                                    & (uhc2_detail.at[7,'OON Copay']!= 'N/A' or uhc2_detail.at[7,'OON Coins']!= 'N/A')
                                                    & (uhc2_detail.at[7,'OON Copay']!= 'No benefit' and uhc2_detail.at[7,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

uhc2_detail.at[8, 'COVERED IN/OON/BOTH'] = np.where((uhc2_detail.at[8,'IN Copay']!= 'N/A' or uhc2_detail.at[8,'IN Coins']!= 'N/A') 
                                                    & (uhc2_detail.at[8,'OON Copay']!= 'N/A' or uhc2_detail.at[8,'OON Coins']!= 'N/A')
                                                    & (uhc2_detail.at[8,'OON Copay']!= 'No benefit' and uhc2_detail.at[8,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

uhc2_detail.at[10, 'COVERED IN/OON/BOTH'] = np.where((uhc2_detail.at[10,'IN Copay']!= 'N/A' or uhc2_detail.at[10,'IN Coins']!= 'N/A') 
                                                    & (uhc2_detail.at[10,'OON Copay']!= 'N/A' or uhc2_detail.at[10,'OON Coins']!= 'N/A')
                                                    & (uhc2_detail.at[10,'OON Copay']!= 'No benefit' and uhc2_detail.at[10,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
uhc2_detail.at[11, 'COVERED IN/OON/BOTH'] = np.where((uhc2_detail.at[11,'IN Copay']!= 'N/A' or uhc2_detail.at[11,'IN Coins']!= 'N/A') 
                                                    & (uhc2_detail.at[11,'OON Copay']!= 'N/A' or uhc2_detail.at[11,'OON Coins']!= 'N/A')
                                                    & (uhc2_detail.at[11,'OON Copay']!= 'No benefit' and uhc2_detail.at[11,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')


path = 'C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/UHC_Example_2_completed.xlsx'

with pd.ExcelWriter(path) as writer:
    uhc2_plan.to_excel(writer, sheet_name='UHC2',index = False)
    uhc2_detail.to_excel(writer,sheet_name = 'UHC2',startrow = 5,index = False)
    
    
    

# creating a pdf file object
pdfObject = open('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Aetna Benefit PPO & Open Access Plan DE-51-100.pdf', 'rb')

# creating a pdf reader object
pdfReader = PdfFileReader(pdfObject)

pdfpagelist = PdfFileReader(pdfObject).getPage(2).extractText().split('\n')



aetna_plan = pd.read_excel('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx', sheet_name='Plan Name')
aetna_plan['Detail']=aetna_plan['Detail'].astype(str)

aetna_plan.at[0, 'Detail'] = pdfpagelist[0][:5]
aetna_plan.at[1, 'Detail'] = pdfpagelist[58][245:275]
aetna_plan.at[2, 'Detail'] = PdfFileReader(pdfObject).getPage(0).extractText().split('\n')[1][:3]
aetna_plan.at[3, 'Detail'] = 'N/A'


aetna_detail = pd.read_excel('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx', sheet_name='Plan Details')
aetna_detail=aetna_detail.astype(str)
for i in range(len(aetna_detail)):
    aetna_detail.at[i, 'IN Copay'] = 'N/A' 
    aetna_detail.at[i, 'IN Coins'] = 'N/A' 
    aetna_detail.at[i, 'IN Coins After Deductible Flag'] = 'N/A' 
    aetna_detail.at[i, 'OON Copay'] = 'N/A' 
    aetna_detail.at[i, 'OON Coins'] = 'N/A' 
    aetna_detail.at[i, 'OON Coins After Deductible Flag'] = 'N/A' 
    aetna_detail.at[i, 'IN DED'] = 'N/A' 
    aetna_detail.at[i, 'IN Max'] = 'N/A' 
    aetna_detail.at[i, 'OON DED'] = 'N/A' 
    aetna_detail.at[i, 'OON Max'] = 'N/A' 
    aetna_detail.at[i, 'COVERED IN/OON/BOTH'] = 'N/A' 

    
for i in (2,9):
    aetna_detail.at[i, 'IN Copay'] = '' 
    aetna_detail.at[i, 'IN Coins'] = '' 
    aetna_detail.at[i, 'IN Coins After Deductible Flag'] = '' 
    aetna_detail.at[i, 'OON Copay'] = '' 
    aetna_detail.at[i, 'OON Coins'] = '' 
    aetna_detail.at[i, 'OON Coins After Deductible Flag'] = '' 
    aetna_detail.at[i, 'IN DED'] = '' 
    aetna_detail.at[i, 'IN Max'] = '' 
    aetna_detail.at[i, 'OON DED'] = '' 
    aetna_detail.at[i, 'OON Max'] = '' 
    aetna_detail.at[i, 'COVERED IN/OON/BOTH'] = ''

aetna_detail.at[0, 'IN DED'] = pdfpagelist[32][:6]
aetna_detail.at[1, 'IN DED'] = pdfpagelist[32][7:13]
aetna_detail.at[0, 'IN Max'] = pdfpagelist[33][:6]
aetna_detail.at[1, 'IN Max'] = pdfpagelist[33][7:13]
aetna_detail.at[0, 'OON DED'] = pdfpagelist[32][14:20]
aetna_detail.at[1, 'OON DED'] = pdfpagelist[32][21:28]
aetna_detail.at[0, 'OON Max'] = pdfpagelist[33][14:21]
aetna_detail.at[1, 'OON Max'] = pdfpagelist[33][22:29]

aetna_detail.at[3, 'IN Coins'] = pdfpagelist[35][:3]
aetna_detail.at[4, 'IN Coins'] = pdfpagelist[36][:3]
aetna_detail.at[5, 'IN Coins'] = pdfpagelist[37][:3]
aetna_detail.at[6, 'IN Coins'] = pdfpagelist[39][:3]
aetna_detail.at[7, 'IN Coins'] = pdfpagelist[40][:3]
aetna_detail.at[8, 'IN Coins'] = pdfpagelist[41][:3]

aetna_detail.at[10, 'IN Coins'] = pdfpagelist[40][:3]
aetna_detail.at[11, 'IN Coins'] = pdfpagelist[42][:3]


aetna_detail.at[3, 'IN Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[4, 'IN Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[5, 'IN Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[6, 'IN Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[7, 'IN Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[8, 'IN Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[10, 'IN Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[11, 'IN Coins After Deductible Flag'] = 'Yes'

aetna_detail.at[3, 'OON Coins'] = pdfpagelist[35][21:24]
aetna_detail.at[4, 'OON Coins'] = pdfpagelist[36][21:24]
aetna_detail.at[5, 'OON Coins'] = pdfpagelist[37][21:24]
aetna_detail.at[6, 'OON Coins'] = pdfpagelist[39][21:24]
aetna_detail.at[7, 'OON Coins'] = pdfpagelist[40][21:24]
aetna_detail.at[8, 'OON Coins'] = pdfpagelist[41][21:24]

aetna_detail.at[10, 'OON Coins'] = pdfpagelist[40][21:24]
aetna_detail.at[11, 'OON Coins'] = pdfpagelist[42][21:24]

aetna_detail.at[3, 'OON Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[4, 'OON Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[5, 'OON Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[6, 'OON Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[7, 'OON Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[8, 'OON Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[10, 'OON Coins After Deductible Flag'] = 'Yes'
aetna_detail.at[11, 'OON Coins After Deductible Flag'] = 'Yes'

aetna_detail.at[3, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail.at[3,'IN Copay']!= 'N/A' or aetna_detail.at[3,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail.at[3,'OON Copay']!= 'N/A' or aetna_detail.at[3,'OON Coins']!= 'N/A')
                                                    & (aetna_detail.at[3,'OON Copay']!= 'No benefit' or aetna_detail.at[3,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
                                                    
aetna_detail.at[4, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail.at[4,'IN Copay']!= 'N/A' or aetna_detail.at[4,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail.at[4,'OON Copay']!= 'N/A' or aetna_detail.at[4,'OON Coins']!= 'N/A')
                                                    & (aetna_detail.at[4,'OON Copay']!= 'No benefit' and aetna_detail.at[4,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

aetna_detail.at[5, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail.at[5,'IN Copay']!= 'N/A' or aetna_detail.at[5,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail.at[5,'OON Copay']!= 'N/A' or aetna_detail.at[5,'OON Coins']!= 'N/A')
                                                    & (aetna_detail.at[5,'OON Copay']!= 'No benefit' and aetna_detail.at[5,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

aetna_detail.at[6, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail.at[6,'IN Copay']!= 'N/A' or aetna_detail.at[6,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail.at[6,'OON Copay']!= 'N/A' or aetna_detail.at[6,'OON Coins']!= 'N/A')
                                                    & (aetna_detail.at[6,'OON Copay']!= 'No benefit' or aetna_detail.at[6,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
                                                    
aetna_detail.at[7, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail.at[7,'IN Copay']!= 'N/A' or aetna_detail.at[7,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail.at[7,'OON Copay']!= 'N/A' or aetna_detail.at[7,'OON Coins']!= 'N/A')
                                                    & (aetna_detail.at[7,'OON Copay']!= 'No benefit' and aetna_detail.at[7,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

aetna_detail.at[8, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail.at[8,'IN Copay']!= 'N/A' or aetna_detail.at[8,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail.at[8,'OON Copay']!= 'N/A' or aetna_detail.at[8,'OON Coins']!= 'N/A')
                                                    & (aetna_detail.at[8,'OON Copay']!= 'No benefit' and aetna_detail.at[8,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')

aetna_detail.at[10, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail.at[10,'IN Copay']!= 'N/A' or aetna_detail.at[10,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail.at[10,'OON Copay']!= 'N/A' or aetna_detail.at[10,'OON Coins']!= 'N/A')
                                                    & (aetna_detail.at[10,'OON Copay']!= 'No benefit' and aetna_detail.at[10,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
aetna_detail.at[11, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail.at[11,'IN Copay']!= 'N/A' or aetna_detail.at[11,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail.at[11,'OON Copay']!= 'N/A' or aetna_detail.at[11,'OON Coins']!= 'N/A')
                                                    & (aetna_detail.at[11,'OON Copay']!= 'No benefit' and aetna_detail.at[11,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')


path = 'C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Aetna_Example_1_completed.xlsx'

with pd.ExcelWriter(path) as writer:
    aetna_plan.to_excel(writer, sheet_name='AETNA',index = False)
    aetna_detail.to_excel(writer,sheet_name = 'AETNA',startrow = 5,index = False)
