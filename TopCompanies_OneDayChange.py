#!/usr/bin/env python
# coding: utf-8

# In[122]:


import requests, openpyxl
from bs4 import BeautifulSoup 


# In[123]:


excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet= excel.active
sheet.title= "Top 100 Strik Dip Today"
print(excel.sheetnames)
sheet.append(['Rank','Company Name','Trade Code','Today Strik/Drip','Country'])


# In[135]:



#url1 = 'https://www.rottentomatoes.com/top/bestofrt/top_100_horror_movies/'
url = 'https://companiesmarketcap.com/'
try:
    source = requests.get(url)
    #print(source)
    source.raise_for_status()

    soup =BeautifulSoup(source.text,'html.parser')
    companies = soup.find('tbody').find_all('tr')
    
    #print(companies)
    
    for company in companies:
        rank = company.find('td').text
        name = company.find('div', class_="company-name").text
        code = company.find('div', class_="company-code").text
        oneDayPercentChange = company.find('td', class_="rh-sm").span.text
        country = company.find('span', class_="responsive-hidden").text
        
        print(rank, name,code,oneDayPercentChange,country)
        sheet.append([rank, name,code,oneDayPercentChange,country])
        
    
except Exception() as e:
    print(e)


# In[ ]:


excel.save("Top One Day Change.xlsx")


# In[ ]:




