#!/usr/bin/env python
# coding: utf-8

# In[4]:


import pandas as pd
import webbrowser 
from datetime import datetime, timedelta
#@from openpyxl import load.workbook


# In[5]:


caminho = r'D:\Onedrive\00_StartUp\25_Doc_da_empresa\Certidões\_Validades.xlsx'
df = pd.read_excel(caminho)
tiposdados = df.dtypes
#print (tiposdados)


# In[6]:


for index , row in df.iterrows():
    data = row['Expira_em']
    doc = row['Documento']
    url = row['Site']

    if pd.isnull(data) or pd.isnull(url):
        continue
    if data < datetime.now():
        print(f"A data de {doc} está vencida.")
        webbrowser.open(url)
        input("Pressione enter após concluir a ação")
print("Terminada a verificação das certidões.")


# In[ ]:




