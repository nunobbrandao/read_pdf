# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import PyPDF2
import re
import pandas as pd
import os 
import glob

#functions 
def most_common(lst,thrs):
    lst_aux = [x for x in lst if x>thrs] 
    return max(set(lst_aux), key=lst_aux.count)

#list of files to be imported
files = glob.glob(os.path.join(os.getcwd(),"*.pdf"))

#from pdfminer.high_level import extract_text
dfs=[]
filenames=[]

#extract the information in each pdf
for file in files:

    print(file)
    sample_pdf = open(file, mode='rb')
    pdfdoc = PyPDF2.PdfFileReader(sample_pdf, strict=False)
    
    #initializing empty lists
    dim_new = []
    dim_hint=[]
    ULS = []
    SLS= []
    drg = []
    naming =[]
    bearing =[]
    
    #extract the information in each page
    for i in range(pdfdoc.numPages):
        page_one= pdfdoc.getPage(i).extractText().replace(",",".").upper()
        #extract data from string
        aux_var = " ".join(re.findall(r'CORROSION PROTECTION ITEM(.*?)\+',page_one, flags = re.DOTALL))
        dim_new.append(" ".join(re.findall(r'\d{2,3}\.\d+|\d{2,3}',aux_var)))
        ULS.append(" ".join(re.findall(r'(?=ULTIMATE)(.*?)\n',page_one)))
        SLS.append(" ".join(re.findall(r'(?=SERVICE)(.*?)\n',page_one)))
        drg.append(" ".join(re.findall(r'193-\d{4}-\d{4}\.\d{0,1}',page_one)))
        naming.append(" ".join(re.findall(r'HM1-P\d\d?\d?',page_one)))
        bearing.append(" ".join(re.findall(r'(?=TYPE)(.*?)(?=\))',page_one, flags = re.DOTALL)))
   
    #extract maximum, most frequent (above a certain threshold), and last from dim_new
    for d in dim_new:
        list_f = list(map(float, d.split()))
        dim_hint.append(" ".join([str(list_f[-1]),str(max(list_f)),str(most_common(list_f,100))]))
    
    #manage numeric data in dataframes
    di = {'SLS':SLS,'ULS':ULS,"Dim_hint":dim_hint,"Dim":dim_new}#,"Type":bearing, "Drg":drg}
    df = pd.DataFrame(di)
    a=[]
    for column in df:
        a.append(df[column].str.split(expand = True))
    df1 = pd.concat(a, axis=1, ignore_index=True)
    df2 = df1.apply(pd.to_numeric, errors='coerce').fillna(df1)
    
    #re-arrange df2
    df2["Name"] = naming 
    df2["Type"] = bearing 
    df2["Drawing"] = drg
    df2["Name_split"]=df2["Name"].str.split()
    df2 = df2.explode("Name_split").set_index('Name_split').sort_index(ascending=True)
    #sorting df2 by numbering of the bearing ID
    df2["sort"] = df2.index.str.extract(r'(\d+)$', expand=False).astype(int)
    df2=df2.sort_values(by='sort', ascending=True).drop('sort',axis=1)
    #save in excel
    filename = os.path.basename(file)
    df2["Filename"]=filename
    filename = os.path.splitext(filename)[0]+'.xlsx'
    df2.to_excel(filename)    
    #create a list of all the datafrane
    dfs.append(df2)

#comparison between different pdf files:
print("************** Comparison Starting:")
comparison = []
for df in dfs:
    comparison.append(df.reset_index(level=0).loc[:,["Name_split","Drawing","Filename"]])
df_comp = pd.concat(comparison)
df_comp = df_comp.loc[df_comp["Name_split"].duplicated(keep=False),:].sort_values('Name_split')
print(df_comp)
df_comp.to_excel("Duplicated_bearings_in_different_pdf.xlsx")

#filter Rendel excel file per PDF
print("************** Filter the RENDEL Excel File:")
#folder = r'C:\Users\n.brandao\OneDrive - INGEROP\1 - Projects\8 - Nuclear project\2 - Bridgwater\4 - Bearing checks\Berings - Rendel raw data'
rendel_excel_file = r'Formated Data HPC-HK2201-U9-HMX-REP-100025 [D].xlsx'
df_r = pd.read_excel(os.path.join(os.getcwd(),rendel_excel_file), sheet_name=1)
df_r.columns = df_r.iloc[1,:]
df_r.iloc[:,0]=df_r.iloc[:,0].str.replace(' ', '')
for df in dfs:
    df_filter=df_r.loc[df_r.iloc[:,0].isin(df.index),:]
    df_filter.to_excel("Rendel"+df.Filename.unique()[0]+".xlsx")
    