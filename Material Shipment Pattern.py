#!/usr/bin/env python
# coding: utf-8

# In[1]:


#Import modules
import pandas as pd
import numpy as py
from mlxtend.frequent_patterns import apriori, association_rules 
from datetime import datetime

#read data excel by sheet
Shipmentreport1=pd.read_excel("Logistic Data Analytics.xlsx", sheet_name=0)
Shipmentreport2=pd.read_excel("Logistic Data Analytics.xlsx", sheet_name=1)
MaterialMD=pd.read_excel("Logistic Data Analytics.xlsx", sheet_name=2)
PlantMD=pd.read_excel("Logistic Data Analytics.xlsx", sheet_name=3)


# In[2]:


#Pre processing
#drop unused row in shipment report 2
x=636468
y=1048555
Shipmentreport2.drop(Shipmentreport2.loc[x:y].index, inplace=True)
Shipmentreport2


# In[3]:


#Join report shipment 1 & 2
Shipmentreport=pd.concat([Shipmentreport1,Shipmentreport2])
Shipmentreport


# In[4]:


#Drop missing values in shipment report
Shipmentreport=Shipmentreport.dropna(how='any', axis=0)
Shipmentreport


# In[5]:


#Aggregate shipment report by referencing material and plant unique id
Aggregate=Shipmentreport.merge(MaterialMD, on='Material', how='left')
Aggregate


# In[6]:


#Create aggregate transaction
transaksi=Aggregate.groupby(['ID Shipment','Material description']).size().reset_index(name='count')
transaksi


# In[7]:


#Create basket transaction
basket=(transaksi.groupby(['ID Shipment','Material description'])['count']
        .sum().unstack().reset_index().fillna(0)
        .set_index('ID Shipment'))
def encode_units(x):
    if x <= 0:
        return 0
    if x >= 1:
        return 1
basket_sets=basket.applymap(encode_units)


# In[8]:


#Create AR
frequent_itemsets=apriori(basket_sets,min_support=0.07,use_colnames=True)
rules=association_rules(frequent_itemsets,metric="lift",min_threshold=1)
rules.sort_values('confidence',ascending=False,inplace=True)
rules


# In[9]:


#rename columns in shipment report
Aggregatechangename=Aggregate.rename(columns={Aggregate.columns[1]: "Qty",Aggregate.columns[6]: "PLANT CODE" })
Aggregatechangename


# In[10]:


#Aggregate shipment report by referencing material and plant unique id
Aggregate2=Aggregatechangename.merge(PlantMD, on='PLANT CODE')
Aggregate2


# In[40]:


#Create aggregate transaction
transaksi2=Aggregate2.groupby(['ID Shipment','Category']).size().reset_index(name='count')
transaksi2


# In[41]:


#Create basket transaction
basket2=(transaksi2.groupby(['ID Shipment','Category'])['count']
        .sum().unstack().reset_index().fillna(0)
        .set_index('ID Shipment'))
def encode_units(x):
    if x <= 0:
        return 0
    if x >= 1:
        return 1
basket_sets2=basket2.applymap(encode_units)


# In[52]:


#Create AR
frequent_itemsets2=apriori(basket_sets2,min_support=0.23,use_colnames=True)
Pattern=association_rules(frequent_itemsets2,metric="lift",min_threshold=1)
Pattern.sort_values('confidence',ascending=False,inplace=True)
Pattern


# In[58]:


#export to excel
Pattern.to_excel(r"E:\Mahendra\KULIAH\SKRIPSI\NEW TOPIC\Data\Category AR.xlsx")
rules.to_excel(r"E:\Mahendra\KULIAH\SKRIPSI\NEW TOPIC\Data\Material AR.xlsx") 


# In[84]:


#create plot
import seaborn as sns
import matplotlib.pyplot as plt

plotmaterial=sns.countplot(x='Material description',data=Aggregate,order=Aggregate['Material description'].value_counts().iloc[:10].index)
plt.xticks(rotation=90)


# In[88]:


#export plot image
fig = plotmaterial.get_figure()
fig.savefig(r'E:\Mahendra\KULIAH\SKRIPSI\NEW TOPIC\Data\Plot Material.png',bbox_inches='tight', dpi=1000)


# In[91]:


#create plot 
plotcategory=sns.countplot(x='Category',data=Aggregate,order=Aggregate['Category'].value_counts().iloc[:5].index)


# In[93]:


#export plot image
fig2 = plotcategory.get_figure()
fig2.savefig(r'E:\Mahendra\KULIAH\SKRIPSI\NEW TOPIC\Data\Plot Category.png')


# In[ ]:




