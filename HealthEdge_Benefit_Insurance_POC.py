"""
HealthEdge Benefit Insurance Template_POC
"""

##Main code
import pandas as pd
import numpy as np
from PyPDF2 import PdfFileReader
from datetime import datetime
from os.path import exists
pd.set_option('display.max_columns', None)

#creating logfile
logfile = pd.DataFrame(columns = ['Time', 'Step', 'Status','Description'])

file_exists_pdf = exists('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Aetna Benefit PPO & Open Access Plan DE-51-100.pdf')
file_exists_notpdf = exists('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Aetna Benefit PPO & Open Access Plan DE-51-100.xlsx')

template_exists_xlsx = exists('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx')
template_exists_notxlsx = exists('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.doc')
template_exists_notxlsx = exists('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.txt')
template_exists_notxlsx = exists('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.pdf')
template_exists_notxlsx = exists('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.docx')
template_exists_notxlsx = exists('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xls')



if file_exists_pdf == False and file_exists_notpdf == False:
    logfile = logfile.append({'Time' : datetime.now(), 'Step' : 'Read benefit file', 'Status' : 'Error','Description': 'Benefit PDF not found'}, ignore_index = True)
elif file_exists_pdf == False and file_exists_notpdf == True:
    logfile = logfile.append({'Time' : datetime.now(), 'Step' : 'Read benefit file', 'Status' : 'Error','Description': 'Benefit PDF is in the wrong format'}, ignore_index = True)
else:
    logfile = logfile.append({'Time' : datetime.now(), 'Step' : 'Read benefit file', 'Status' : 'Complete','Description': 'Benefit PDF load completed'}, ignore_index = True)


if template_exists_xlsx == False and template_exists_notxlsx == False:
    logfile = logfile.append({'Time' : datetime.now(), 'Step' : 'Read in template', 'Status' : 'Error','Description': 'Template not found'}, ignore_index = True)
elif template_exists_xlsx == False and template_exists_notxlsx == True:
    logfile = logfile.append({'Time' : datetime.now(), 'Step' : 'Read in template', 'Status' : 'Error','Description': 'Template is in teh wrong format'}, ignore_index = True)
else:
    logfile = logfile.append({'Time' : datetime.now(), 'Step' : 'Read in template', 'Status' : 'Complete','Description': 'Template load completed'}, ignore_index = True)

logfile.to_csv('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/logfile.txt', index=None, sep=' ', mode='w+')


try:  
    pdfObject = open('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Aetna Benefit PPO & Open Access Plan DE-51-100.pdf', 'rb')
    pdfReader = PdfFileReader(pdfObject)
    pdfpagelist_1 = PdfFileReader(pdfObject).getPage(0).extractText().split('\n')
except:
    logfile.to_csv('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/logfile.txt', index=None, sep=' ', mode='w+')
else:
    logfile = logfile.append({'Time' : datetime.now(), 'Step' : 'Complete template', 'Status' : 'Complete','Description': 'Completed template written'}, ignore_index = True)
    aetna_plan_1 = pd.read_excel('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx', sheet_name='Plan Name')
    aetna_plan_1['Detail']=aetna_plan_1['Detail'].astype(str)
    aetna_plan_1.at[0, 'Detail'] = pdfpagelist_1[0][:5]
    aetna_plan_1.at[1, 'Detail'] = pdfpagelist_1[45][11:35]
    aetna_plan_1.at[2, 'Detail'] = PdfFileReader(pdfObject).getPage(0).extractText().split('\n')[1][:3]
    aetna_plan_1.at[3, 'Detail'] = 'N/A'
    aetna_detail_1 = pd.read_excel('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx', sheet_name='Plan Details')
    aetna_detail_1=aetna_detail_1.astype(str)
    for i in range(len(aetna_detail_1)):
        aetna_detail_1.at[i, 'IN Copay'] = 'N/A' 
        aetna_detail_1.at[i, 'IN Coins'] = 'N/A' 
        aetna_detail_1.at[i, 'IN Coins After Deductible Flag'] = 'N/A' 
        aetna_detail_1.at[i, 'OON Copay'] = 'N/A' 
        aetna_detail_1.at[i, 'OON Coins'] = 'N/A' 
        aetna_detail_1.at[i, 'OON Coins After Deductible Flag'] = 'N/A' 
        aetna_detail_1.at[i, 'IN DED'] = 'N/A' 
        aetna_detail_1.at[i, 'IN Max'] = 'N/A' 
        aetna_detail_1.at[i, 'OON DED'] = 'N/A' 
        aetna_detail_1.at[i, 'OON Max'] = 'N/A' 
        aetna_detail_1.at[i, 'COVERED IN/OON/BOTH'] = 'N/A' 
    for i in (2,9):
         aetna_detail_1.at[i, 'IN Copay'] = '' 
         aetna_detail_1.at[i, 'IN Coins'] = '' 
         aetna_detail_1.at[i, 'IN Coins After Deductible Flag'] = '' 
         aetna_detail_1.at[i, 'OON Copay'] = '' 
         aetna_detail_1.at[i, 'OON Coins'] = '' 
         aetna_detail_1.at[i, 'OON Coins After Deductible Flag'] = '' 
         aetna_detail_1.at[i, 'IN DED'] = '' 
         aetna_detail_1.at[i, 'IN Max'] = '' 
         aetna_detail_1.at[i, 'OON DED'] = '' 
         aetna_detail_1.at[i, 'OON Max'] = '' 
         aetna_detail_1.at[i, 'COVERED IN/OON/BOTH'] = ''
         
         aetna_detail_1.at[0, 'IN DED'] = pdfpagelist_1[5][31:37]
         aetna_detail_1.at[1, 'IN DED'] = pdfpagelist_1[5][38:44]
         aetna_detail_1.at[0, 'IN Max'] = pdfpagelist_1[6][40:46]
         aetna_detail_1.at[1, 'IN Max'] = pdfpagelist_1[6][47:54]
         aetna_detail_1.at[0, 'OON DED'] = pdfpagelist_1[5][45:51]
         aetna_detail_1.at[1, 'OON DED'] = pdfpagelist_1[5][52:59]
         aetna_detail_1.at[0, 'OON Max'] = pdfpagelist_1[6][55:62]
         aetna_detail_1.at[1, 'OON Max'] = pdfpagelist_1[6][63:70]
         
         aetna_detail_1.at[3, 'IN Copay'] = pdfpagelist_1[8][36:39]
         aetna_detail_1.at[4, 'IN Copay'] = pdfpagelist_1[9][24:27]
         aetna_detail_1.at[5, 'IN Copay'] = '$0'
         aetna_detail_1.at[6, 'IN Copay'] = pdfpagelist_1[14][26:30]
         aetna_detail_1.at[7, 'IN Copay'] = pdfpagelist_1[9][24:27]
         aetna_detail_1.at[8, 'IN Copay'] = pdfpagelist_1[8][36:39]
         
         aetna_detail_1.at[10, 'IN Copay'] = '$0'
         aetna_detail_1.at[11, 'IN Copay'] = pdfpagelist_1[14][26:30]
         
         aetna_detail_1.at[3, 'IN Coins After Deductible Flag'] = 'No'
         aetna_detail_1.at[4, 'IN Coins After Deductible Flag'] = 'No'
         aetna_detail_1.at[5, 'IN Coins After Deductible Flag'] = 'No'
         aetna_detail_1.at[6, 'IN Coins After Deductible Flag'] = 'No'
         aetna_detail_1.at[7, 'IN Coins After Deductible Flag'] = 'No'
         aetna_detail_1.at[8, 'IN Coins After Deductible Flag'] = 'No'
         aetna_detail_1.at[10, 'IN Coins After Deductible Flag'] = 'Yes'
         aetna_detail_1.at[11, 'IN Coins After Deductible Flag'] = 'No'
         
         
         aetna_detail_1.at[3, 'OON Coins'] = pdfpagelist_1[8][65:68]
         aetna_detail_1.at[4, 'OON Coins'] = pdfpagelist_1[8][65:68]
         aetna_detail_1.at[5, 'OON Coins'] = pdfpagelist_1[8][65:68]
         aetna_detail_1.at[6, 'OON Coins'] = pdfpagelist_1[8][65:68]
         aetna_detail_1.at[7, 'OON Coins'] = pdfpagelist_1[8][65:68]
         aetna_detail_1.at[8, 'OON Coins'] = pdfpagelist_1[8][65:68]
         
         aetna_detail_1.at[10, 'OON Coins'] = pdfpagelist_1[8][65:68]
         aetna_detail_1.at[11, 'OON Coins'] = pdfpagelist_1[8][65:68]
         
         aetna_detail_1.at[3, 'OON Coins After Deductible Flag'] = 'Yes'
         aetna_detail_1.at[4, 'OON Coins After Deductible Flag'] = 'Yes'
         aetna_detail_1.at[5, 'OON Coins After Deductible Flag'] = 'Yes'
         aetna_detail_1.at[6, 'OON Coins After Deductible Flag'] = 'Yes'
         aetna_detail_1.at[7, 'OON Coins After Deductible Flag'] = 'Yes'
         aetna_detail_1.at[8, 'OON Coins After Deductible Flag'] = 'Yes'
         aetna_detail_1.at[10, 'OON Coins After Deductible Flag'] = 'Yes'
         aetna_detail_1.at[11, 'OON Coins After Deductible Flag'] = 'Yes'
         
         aetna_detail_1.at[3, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_1.at[3,'IN Copay']!= 'N/A' or aetna_detail_1.at[3,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_1.at[3,'OON Copay']!= 'N/A' or aetna_detail_1.at[3,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_1.at[3,'OON Copay']!= 'No benefit' or aetna_detail_1.at[3,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
         
         aetna_detail_1.at[4, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_1.at[4,'IN Copay']!= 'N/A' or aetna_detail_1.at[4,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_1.at[4,'OON Copay']!= 'N/A' or aetna_detail_1.at[4,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_1.at[4,'OON Copay']!= 'No benefit' and aetna_detail_1.at[4,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
         
         aetna_detail_1.at[5, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_1.at[5,'IN Copay']!= 'N/A' or aetna_detail_1.at[5,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_1.at[5,'OON Copay']!= 'N/A' or aetna_detail_1.at[5,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_1.at[5,'OON Copay']!= 'No benefit' and aetna_detail_1.at[5,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
         
         aetna_detail_1.at[6, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_1.at[6,'IN Copay']!= 'N/A' or aetna_detail_1.at[6,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_1.at[6,'OON Copay']!= 'N/A' or aetna_detail_1.at[6,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_1.at[6,'OON Copay']!= 'No benefit' or aetna_detail_1.at[6,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
         
         aetna_detail_1.at[7, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_1.at[7,'IN Copay']!= 'N/A' or aetna_detail_1.at[7,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_1.at[7,'OON Copay']!= 'N/A' or aetna_detail_1.at[7,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_1.at[7,'OON Copay']!= 'No benefit' and aetna_detail_1.at[7,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
         
         aetna_detail_1.at[8, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_1.at[8,'IN Copay']!= 'N/A' or aetna_detail_1.at[8,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_1.at[8,'OON Copay']!= 'N/A' or aetna_detail_1.at[8,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_1.at[8,'OON Copay']!= 'No benefit' and aetna_detail_1.at[8,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
         
         aetna_detail_1.at[10, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_1.at[10,'IN Copay']!= 'N/A' or aetna_detail_1.at[10,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_1.at[10,'OON Copay']!= 'N/A' or aetna_detail_1.at[10,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_1.at[10,'OON Copay']!= 'No benefit' and aetna_detail_1.at[10,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
         aetna_detail_1.at[11, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_1.at[11,'IN Copay']!= 'N/A' or aetna_detail_1.at[11,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_1.at[11,'OON Copay']!= 'N/A' or aetna_detail_1.at[11,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_1.at[11,'OON Copay']!= 'No benefit' and aetna_detail_1.at[11,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
         
         
         path = 'C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Aetna_PPO_examples.xlsx'
         
         
         
         pdfObject = open('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Aetna Benefit PPO & Open Access Plan DE-51-100.pdf', 'rb')
         pdfReader = PdfFileReader(pdfObject)
         pdfpagelist_2 = PdfFileReader(pdfObject).getPage(2).extractText().split('\n')
         
         aetna_plan_2 = pd.read_excel('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx', sheet_name='Plan Name')
         aetna_plan_2['Detail']=aetna_plan_2['Detail'].astype(str)
         aetna_plan_2.at[0, 'Detail'] = pdfpagelist_2[0][:5]
         aetna_plan_2.at[1, 'Detail'] = pdfpagelist_2[58][245:275]
         aetna_plan_2.at[2, 'Detail'] = PdfFileReader(pdfObject).getPage(0).extractText().split('\n')[1][:3]
         aetna_plan_2.at[3, 'Detail'] = 'N/A'
         aetna_detail_2 = pd.read_excel('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Insurance Template.xlsx', sheet_name='Plan Details')
         aetna_detail_2=aetna_detail_2.astype(str)
         for i in range(len(aetna_detail_2)):
             aetna_detail_2.at[i, 'IN Copay'] = 'N/A' 
             aetna_detail_2.at[i, 'IN Coins'] = 'N/A' 
             aetna_detail_2.at[i, 'IN Coins After Deductible Flag'] = 'N/A' 
             aetna_detail_2.at[i, 'OON Copay'] = 'N/A' 
             aetna_detail_2.at[i, 'OON Coins'] = 'N/A' 
             aetna_detail_2.at[i, 'OON Coins After Deductible Flag'] = 'N/A' 
             aetna_detail_2.at[i, 'IN DED'] = 'N/A' 
             aetna_detail_2.at[i, 'IN Max'] = 'N/A' 
             aetna_detail_2.at[i, 'OON DED'] = 'N/A' 
             aetna_detail_2.at[i, 'OON Max'] = 'N/A' 
             aetna_detail_2.at[i, 'COVERED IN/OON/BOTH'] = 'N/A' 
         for i in (2,9):
             aetna_detail_2.at[i, 'IN Copay'] = '' 
             aetna_detail_2.at[i, 'IN Coins'] = '' 
             aetna_detail_2.at[i, 'IN Coins After Deductible Flag'] = '' 
             aetna_detail_2.at[i, 'OON Copay'] = '' 
             aetna_detail_2.at[i, 'OON Coins'] = '' 
             aetna_detail_2.at[i, 'OON Coins After Deductible Flag'] = '' 
             aetna_detail_2.at[i, 'IN DED'] = '' 
             aetna_detail_2.at[i, 'IN Max'] = '' 
             aetna_detail_2.at[i, 'OON DED'] = '' 
             aetna_detail_2.at[i, 'OON Max'] = '' 
             aetna_detail_2.at[i, 'COVERED IN/OON/BOTH'] = ''
             
             aetna_detail_2.at[0, 'IN DED'] = pdfpagelist_2[32][:6]
             aetna_detail_2.at[1, 'IN DED'] = pdfpagelist_2[32][7:13]
             aetna_detail_2.at[0, 'IN Max'] = pdfpagelist_2[33][:6]
             aetna_detail_2.at[1, 'IN Max'] = pdfpagelist_2[33][7:13]
             aetna_detail_2.at[0, 'OON DED'] = pdfpagelist_2[32][14:20]
             aetna_detail_2.at[1, 'OON DED'] = pdfpagelist_2[32][21:28]
             aetna_detail_2.at[0, 'OON Max'] = pdfpagelist_2[33][14:21]
             aetna_detail_2.at[1, 'OON Max'] = pdfpagelist_2[33][22:29]
             
             aetna_detail_2.at[3, 'IN Coins'] = pdfpagelist_2[35][:3]
             aetna_detail_2.at[4, 'IN Coins'] = pdfpagelist_2[36][:3]
             aetna_detail_2.at[5, 'IN Coins'] = pdfpagelist_2[37][:3]
             aetna_detail_2.at[6, 'IN Coins'] = pdfpagelist_2[39][:3]
             aetna_detail_2.at[7, 'IN Coins'] = pdfpagelist_2[40][:3]
             aetna_detail_2.at[8, 'IN Coins'] = pdfpagelist_2[41][:3]
             
             aetna_detail_2.at[10, 'IN Coins'] = pdfpagelist_2[40][:3]
             aetna_detail_2.at[11, 'IN Coins'] = pdfpagelist_2[42][:3]
             
             aetna_detail_2.at[3, 'IN Coins After Deductible Flag'] = 'Yes'
             aetna_detail_2.at[4, 'IN Coins After Deductible Flag'] = 'Yes'
             aetna_detail_2.at[5, 'IN Coins After Deductible Flag'] = 'Yes'
             aetna_detail_2.at[6, 'IN Coins After Deductible Flag'] = 'Yes'
             aetna_detail_2.at[7, 'IN Coins After Deductible Flag'] = 'Yes'
             aetna_detail_2.at[8, 'IN Coins After Deductible Flag'] = 'Yes'
             aetna_detail_2.at[10, 'IN Coins After Deductible Flag'] = 'Yes'
             aetna_detail_2.at[11, 'IN Coins After Deductible Flag'] = 'Yes'
             
             aetna_detail_2.at[3, 'OON Coins'] = pdfpagelist_2[35][21:24]
             aetna_detail_2.at[4, 'OON Coins'] = pdfpagelist_2[36][21:24]
             aetna_detail_2.at[5, 'OON Coins'] = pdfpagelist_2[37][21:24]
             aetna_detail_2.at[6, 'OON Coins'] = pdfpagelist_2[39][21:24]
             aetna_detail_2.at[7, 'OON Coins'] = pdfpagelist_2[40][21:24]
             aetna_detail_2.at[8, 'OON Coins'] = pdfpagelist_2[41][21:24]
             
             aetna_detail_2.at[10, 'OON Coins'] = pdfpagelist_2[40][21:24]
             aetna_detail_2.at[11, 'OON Coins'] = pdfpagelist_2[42][21:24]
             
             aetna_detail_2.at[3, 'OON Coins After Deductible Flag'] = 'No'
             aetna_detail_2.at[4, 'OON Coins After Deductible Flag'] = 'No'
             aetna_detail_2.at[5, 'OON Coins After Deductible Flag'] = 'No'
             aetna_detail_2.at[6, 'OON Coins After Deductible Flag'] = 'Yes'
             aetna_detail_2.at[7, 'OON Coins After Deductible Flag'] = 'Yes'
             aetna_detail_2.at[8, 'OON Coins After Deductible Flag'] = 'Yes'
             aetna_detail_2.at[10, 'OON Coins After Deductible Flag'] = 'Yes'
             aetna_detail_2.at[11, 'OON Coins After Deductible Flag'] = 'Yes'
             
             aetna_detail_2.at[3, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_2.at[3,'IN Copay']!= 'N/A' or aetna_detail_2.at[3,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_2.at[3,'OON Copay']!= 'N/A' or aetna_detail_2.at[3,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_2.at[3,'OON Copay']!= 'No benefit' or aetna_detail_2.at[3,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
             
             aetna_detail_2.at[4, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_2.at[4,'IN Copay']!= 'N/A' or aetna_detail_2.at[4,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_2.at[4,'OON Copay']!= 'N/A' or aetna_detail_2.at[4,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_2.at[4,'OON Copay']!= 'No benefit' and aetna_detail_2.at[4,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
             
             aetna_detail_2.at[5, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_2.at[5,'IN Copay']!= 'N/A' or aetna_detail_2.at[5,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_2.at[5,'OON Copay']!= 'N/A' or aetna_detail_2.at[5,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_2.at[5,'OON Copay']!= 'No benefit' and aetna_detail_2.at[5,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
             
             aetna_detail_2.at[6, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_2.at[6,'IN Copay']!= 'N/A' or aetna_detail_2.at[6,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_2.at[6,'OON Copay']!= 'N/A' or aetna_detail_2.at[6,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_2.at[6,'OON Copay']!= 'No benefit' or aetna_detail_2.at[6,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
             
             aetna_detail_2.at[7, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_2.at[7,'IN Copay']!= 'N/A' or aetna_detail_2.at[7,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_2.at[7,'OON Copay']!= 'N/A' or aetna_detail_2.at[7,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_2.at[7,'OON Copay']!= 'No benefit' and aetna_detail_2.at[7,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
             
             aetna_detail_2.at[8, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_2.at[8,'IN Copay']!= 'N/A' or aetna_detail_2.at[8,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_2.at[8,'OON Copay']!= 'N/A' or aetna_detail_2.at[8,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_2.at[8,'OON Copay']!= 'No benefit' and aetna_detail_2.at[8,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
             
             aetna_detail_2.at[10, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_2.at[10,'IN Copay']!= 'N/A' or aetna_detail_2.at[10,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_2.at[10,'OON Copay']!= 'N/A' or aetna_detail_2.at[10,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_2.at[10,'OON Copay']!= 'No benefit' and aetna_detail_2.at[10,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
             aetna_detail_2.at[11, 'COVERED IN/OON/BOTH'] = np.where((aetna_detail_2.at[11,'IN Copay']!= 'N/A' or aetna_detail_2.at[11,'IN Coins']!= 'N/A') 
                                                    & (aetna_detail_2.at[11,'OON Copay']!= 'N/A' or aetna_detail_2.at[11,'OON Coins']!= 'N/A')
                                                    & (aetna_detail_2.at[11,'OON Copay']!= 'No benefit' and aetna_detail_2.at[11,'OON Coins']!= 'No benefit')
                                                    ,'Both','IN Only')
             
             
             
             path = 'C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/Aetna_PPO_examples.xlsx'
             
             with pd.ExcelWriter(path) as writer:
                 aetna_plan_1.to_excel(writer, sheet_name='AETNA EXAMPLE 1',index = False)
                 aetna_detail_1.to_excel(writer,sheet_name = 'AETNA EXAMPLE 1',startrow = 5,index = False)
                 aetna_plan_2.to_excel(writer, sheet_name='AETNA EXAMPLE 2',index = False)
                 aetna_detail_2.to_excel(writer,sheet_name = 'AETNA EXAMPLE 2',startrow = 5,index = False)


logfile.to_csv('C:/Users/yonid/Downloads/HealthEdge Insurance Benefit POC Folder/logfile.txt', index=None, sep=' ', mode='w+')