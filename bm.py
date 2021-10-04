# -*- coding: utf-8 -*-
"""
Created on Tue Sep 28 13:53:47 2021

@author: behzo
"""

import pandas as pd
import datetime


# HAFTA RAQAMI!!!
weeknumber = int(datetime.datetime.now().strftime("%V"))
# =============================================================================
#SKU lar ro'yxati
skulistfile = pd.read_excel('C:\\BM\\SKU list.xlsx')
# =============================================================================
weeklydata = pd.read_excel('C:\\Python\\template excel\\weekly40data.xlsx')
###


tashkent = ['A10R','A12L','A18L','A23','A30R','A31L','A7L','A8','A9L','Abu sahiy dukon',
'Abu sahiy dukon Official','Abu Sahiy Store','B13R','B33L','B4R','B9L',
'Eldor Aka Malika Store','Eldor Aka Malika Store Official','Eldor Uzb.Sklad',
'Eldor Uzb.Sklad Official','Main stock uzb','Main stock uzb Official',
'Malika Main','Malika Murod','Malika Murod Official','Malika Sklad',
'Samarqand','Stock 2','Vivo','Xoji Abu Sahiy','Xoji Abu Sahiy Official']

namangan = ['Namangan']
andijon = ['Andijon']
fargona = ['Fargona']
qoqon = ['Quqon Showroom', 'Quqon Sklad']
sirdaryo = ['Sirdaryo', "Sirdaryo Official"]
jizzax = ['Jizzax', "Jizzax Official"]
navoiy = ['Navoyi', "Navoyi Official"]
samarqand = ['Kattakurgon Official','Kattakurgon', 'Samarqand']
qashqadaryo = ['Qashqadaryo Official', 'Qashqadaryo']
surxondaryo = ['Surxandaryo Official', 'Surxandaryo']
buxoro = ['Buxoro Official', 'Buxoro']
xorazm = ['Xorazm Official', 'Xorazm']
nukus = ['Nukus Official', 'Nukus']

regions = []
for outlet in weeklydata['Outlet Name'].tolist():
    if outlet in tashkent:
        regions.append('Tashkent')
    elif outlet in namangan:
        regions.append('Namangan')
    elif outlet in andijon:
        regions.append('Andijon')
    elif outlet in fargona:
        regions.append('Fargona')
    elif outlet in qoqon:
        regions.append('Qoqon')
    elif outlet in sirdaryo:
        regions.append('Sirdaryo')
    elif outlet in jizzax:
        regions.append('Jizzax')
    elif outlet in navoiy:
        regions.append('Navoiy')
    elif outlet in samarqand:
        regions.append('Samarqand')
    elif outlet in qashqadaryo:
        regions.append('Qashqadaryo')
    elif outlet in surxondaryo:
        regions.append('Surxondaryo')        
    elif outlet in buxoro:
        regions.append('Buxoro')
    elif outlet in xorazm:
        regions.append('Xorazm')
    elif outlet in nukus:
        regions.append('Nukus')
    else:
        regions.append('Nope')        

weeklydata['Regions'] = regions

weeklydata["Week"] = weeklydata["Week"].astype("category")
weeklydata["Outlet Name"] = weeklydata["Outlet Name"].astype("category")
#weeklydata["Item Group"] = weeklydata["Item Group"].astype("category")
#weeklydata["Week"].cat.set_categories(["38","39"],inplace=True)

#faqat 39-haftani qoldirish
weeklydata = weeklydata[weeklydata["Week"]==weeknumber]
pivot = pd.pivot_table(weeklydata, index = ["Item Name","Regions"], values = ["P", "S", "I"], aggfunc = 'sum', fill_value=0)
pivot.reset_index(inplace = True)



itemgroup = []
sku = []
for itemname in pivot["Item Name"].tolist():
    itemgroup.append(weeklydata["Item Group"].tolist()[weeklydata["Item Name"].tolist().index(itemname)])
  
pivot["Item Group"] = itemgroup

pivot['Item Name'] = pivot["Item Name"].str.replace(" PCT","")
pivot['Item Name'] = pivot["Item Name"].str.replace(" GLOBAL","")
pivot['Item Name'] = pivot["Item Name"].str.replace("GLOBAL","")
pivot['Item Name'] = pivot["Item Name"].str.replace("GRAY ","GRAY")
pivot['Item Name'] = pivot["Item Name"].str.replace("GREY","GRAY")

for itemname in pivot["Item Name"].tolist():
    if itemname in skulistfile["INPUT_MODEL"].tolist():
        sku.append(skulistfile["SKU"].tolist()[skulistfile["INPUT_MODEL"].tolist().index(itemname)])
    else:
        sku.append('Not found')
pivot["SKU"] = sku

neworder = ['SKU','Item Name', 'Regions', 'P', 'S', 'I', 'Item Group']
pivot = pivot.reindex(columns=neworder)

# filtrlab olish region boyicha model boyicha
umumiy_samsung = pivot.query("`Item Group`=='SAMSUNG PCT'")
umumiy_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT']")
umumiy_vivo = pivot.query("`Item Group`=='VIVO'")

tashkent_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Tashkent'")
tashkent_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Tashkent'")
tashkent_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Tashkent'")

namangan_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Namangan'")
namangan_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Namangan'")
namangan_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Namangan'")

andijon_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Andijon'")
andijon_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Andijon'")
andijon_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Andijon'")

fargona_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Fargona'")
fargona_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Fargona'")
fargona_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Fargona'")

qoqon_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Qoqon'")
qoqon_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Qoqon'")
qoqon_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Qoqon'")

sirdaryo_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Sirdaryo'")
sirdaryo_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Sirdaryo'")
sirdaryo_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Sirdaryo'")

samarqand_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Samarqand'")
samarqand_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Samarqand'")
samarqand_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Samarqand'")

jizzax_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Jizzax'")
jizzax_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Jizzax'")
jizzax_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Jizzax'")

qashqadaryo_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Qashqadaryo'")
qashqadaryo_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Qashqadaryo'")
qashqadaryo_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Qashqadaryo'")

surxondaryo_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Surxondaryo'")
surxondaryo_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Surxondaryo'")
surxondaryo_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Surxondaryo'")

navoiy_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Navoiy'")
navoiy_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Navoiy'")
navoiy_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Navoiy'")

buxoro_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Buxoro'")
buxoro_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Buxoro'")
buxoro_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Buxoro'")

xorazm_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Xorazm'")
xorazm_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Xorazm'")
xorazm_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Xorazm'")

nukus_samsung = pivot.query("`Item Group`=='SAMSUNG PCT' and Regions == 'Nukus'")
nukus_xiaomi = pivot.query("`Item Group`== ['XIAOMI', 'XIAOMI PCT'] and Regions == 'Nukus'")
nukus_vivo = pivot.query("`Item Group`=='VIVO' and Regions == 'Nukus'")

#excelga saqlash
samsung = pd.ExcelWriter(f'samsung{weeknumber}.xlsx')
for i in ['umumiy_samsung','tashkent_samsung','namangan_samsung','andijon_samsung',
          'fargona_samsung', 'qoqon_samsung', 'sirdaryo_samsung',
          'jizzax_samsung','samarqand_samsung', 'navoiy_samsung', 'qashqadaryo_samsung', 
          'surxondaryo_samsung',"buxoro_samsung", 'xorazm_samsung', 'nukus_samsung']:
       eval(i).to_excel(excel_writer=samsung,sheet_name=i,index=False)
samsung.save()
samsung.close()


    
xiaomi = pd.ExcelWriter(f'xiaomi{weeknumber}.xlsx')

for i in ['umumiy_xiaomi','tashkent_xiaomi','namangan_xiaomi','andijon_xiaomi',
          'fargona_xiaomi', 'qoqon_xiaomi', 'sirdaryo_xiaomi',
          'jizzax_xiaomi','samarqand_xiaomi', 'navoiy_xiaomi', 'qashqadaryo_xiaomi', 
          'surxondaryo_xiaomi',"buxoro_xiaomi", 'xorazm_xiaomi', 'nukus_xiaomi']:
       eval(i).to_excel(excel_writer=xiaomi,sheet_name=i,index=False)
xiaomi.save()
xiaomi.close()


vivo = pd.ExcelWriter(f'vivo{weeknumber}.xlsx')
for i in ['umumiy_vivo','tashkent_vivo','namangan_vivo','andijon_vivo',
          'fargona_vivo', 'qoqon_vivo', 'sirdaryo_vivo',
          'jizzax_vivo','samarqand_vivo', 'navoiy_vivo', 'qashqadaryo_vivo', 
          'surxondaryo_vivo',"buxoro_vivo", 'xorazm_vivo', 'nukus_vivo']:
       eval(i).to_excel(excel_writer=vivo,sheet_name=i,index=False)
vivo.save()
vivo.close()
