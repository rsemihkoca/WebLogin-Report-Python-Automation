# -*- coding: utf-8 -*-
"""
Created on Tue Oct 19 10:54:56 2021

@author: Rıza Semih Koca
"""

from timeit import default_timer as timer
import pandas as pd
import sys


import os
dir_path = os.path.dirname(os.path.abspath('__file__'))


from datetime import datetime

now = datetime.now().strftime("%H.%M %d%m%Y")

ftl_report_name=''
Login_info_name=''
cust_list_name=''
one_monthbefore_name=''
two_monthbefore_name=''

import tkinter as tk
from tkinter.filedialog import askopenfilename

root = tk.Tk()
root.iconify()


_= askopenfilename(initialdir=dir_path,title="Select Customer List")
cust_list_name=str(_).replace("/","\\")
print(cust_list_name)

def dosyalari_alma():
    global ftl_report_name
    global Login_info_name
    global cust_list_name
    global one_monthbefore_name
    global two_monthbefore_name

    if cust_list_name=='':
        root.destroy()
        return sys.exit('IPTAL edildi')

    else:

        _= askopenfilename(initialdir=dir_path,title="Select FTL List")
        ftl_report_name=str(_).replace("/","\\")
        print(ftl_report_name)

        if ftl_report_name=='':
            root.destroy()
            return sys.exit('IPTAL edildi')
        else:

            _= askopenfilename(initialdir=dir_path,title="Select Login List")
            Login_info_name=str(_).replace("/","\\")
            print(Login_info_name)

            if Login_info_name=='':
                root.destroy()
                return sys.exit('IPTAL edildi')
            else:
                _= askopenfilename(initialdir=dir_path,title="Select 1monthbefore")
                one_monthbefore_name=str(_).replace("/","\\")
                print(one_monthbefore_name)

                if one_monthbefore_name=='':
                    root.destroy()
                    return sys.exit('IPTAL edildi')
                else:

                    _ = askopenfilename(initialdir=dir_path,title="Select 2monthbefore")
                    two_monthbefore_name=str(_).replace("/","\\")
                    print(two_monthbefore_name)
                    if two_monthbefore_name=='':
                        root.destroy()
                        return sys.exit('IPTAL edildi')

dosyalari_alma()
start = timer()
print("Dosyalar alındı","\t",round(timer()-start,1))
root.destroy()

web_login_aug = pd.read_excel(one_monthbefore_name,sheet_name='Data',usecols=["Customer Code","Ağustos Login Adet","Eylül Login Adet"],dtype={'Customer Code': str}, engine='openpyxl')

print("Completed",'Weblogin Eylül Final.xlsx',"\t",round(timer()-start,1))

ftl_report = pd.read_excel(ftl_report_name,usecols=["Bat_Id__c","TAPDK_ID__c","Mobil_Login_Durumu__c"],dtype={'Bat_Id__c': str, 'TAPDK_ID__c': str, 'Mobil_Login_Durumu__c': str })
ftl_report['TAPDK_ID__c']='Evet'
ftl_report.rename(columns={'Bat_Id__c':'Customer Code', 'TAPDK_ID__c':'Web login', 'Mobil_Login_Durumu__c':'Mobil login'},inplace=True)

print("Completed",ftl_report_name,"\t",round(timer()-start,1))

cust_list = pd.read_excel(cust_list_name, dtype={'Customer Code': str},engine="openpyxl")
# usecols=["Division","Region","Branch","City","Customer Code","KD Cluster"]
print("Completed",cust_list_name,"\t",round(timer()-start,1))

LoginInformation = pd.read_excel(Login_info_name, dtype={'BatId__c':str,'Number_of_Login__c':int})
LoginInformation.rename(columns={'BatId__c':'Customer Code',LoginInformation.columns[1]:'Ekim Login Adet'},inplace=True)
LoginInformation = LoginInformation.groupby(['Customer Code']).agg('sum')# sum up duplicated custcode's login number
LoginInformation = LoginInformation.reset_index(drop=False)

print("Completed1 Duplication Check Has Started:",Login_info_name,"\t",round(timer()-start,1))

def check_duplicates(*kwargs):
    print('Do we have duplicated row for imported datas:')
    for _ in range(len(kwargs)):
        print(kwargs[_]['Customer Code'].duplicated().any())

check_duplicates(ftl_report,LoginInformation,cust_list)

print("Completed2 Data Sheet Creation:","\t",round(timer()-start,1))

weblogin_pivot=ftl_report[['Customer Code','Web login']]
moblogin_pivot=ftl_report[['Customer Code','Mobil login']]
Log_Info_eyl=LoginInformation[['Customer Code','Ekim Login Adet']]

Log_Info_eyl_webregular=LoginInformation[['Customer Code','Ekim Login Adet']]
Log_Info_eyl_webregular['Ekim Web Regular']=LoginInformation['Ekim Login Adet'].transform(lambda x: 1 if x>0 else 0)
del Log_Info_eyl_webregular['Ekim Login Adet']

Log_Info_tem=web_login_aug[['Customer Code','Ağustos Login Adet']]
Log_Info_aug=web_login_aug[['Customer Code','Eylül Login Adet']]


Heading_list=[
    {"Name":'Web Login',"Pivot":weblogin_pivot,"Fillna":'Hayır'},
    {"Name":'Mobil Login',"Pivot":moblogin_pivot,"Fillna":'Hayır'},
    {"Name":'Ekim Login Adet',"Pivot":Log_Info_eyl,"Fillna":0},
    {"Name":'Ekim Web Regular',"Pivot":Log_Info_eyl_webregular,"Fillna":0},
    {"Name":'Eylül Login Adet',"Pivot":Log_Info_aug,"Fillna":0},
    {"Name":'Ağustos Login Adet',"Pivot":Log_Info_tem,"Fillna":0}

]

def Data_leftjoin(left,via,Heading_list):

    Left_join=left
    for dic in Heading_list:
        Left_join = pd.merge(Left_join, dic["Pivot"], on=via, how ='left').fillna(dic["Fillna"])
        Left_join = Left_join.rename({dic["Pivot"].columns[-1]:  dic["Name"]}, axis='columns')# tmr planner yerine column name yap hepsi tmr değil

    return Left_join


z=Data_leftjoin(cust_list,'Customer Code',Heading_list)

print("Completed3 Manuel handling:","\t",round(timer()-start,1))

def manuel_duzeltme():
    z.loc[(z['Ekim Login Adet']>0) & (z['Web Login']=="Hayır"),'Web Login']='Evet'
    z.loc[(z['Eylül Login Adet']>0) & (z['Web Login']=="Hayır"),'Web Login']='Evet'
    z.loc[(z['Ağustos Login Adet']>0) & (z['Web Login']=="Hayır"),'Web Login']='Evet'


manuel_duzeltme()

print("Completed4 SUM Sheet Creation:","\t",round(timer()-start,1))

z_index_columns= z.pivot_table(index=['Division','Region']).iloc[:,[]].reset_index(drop=False)

reg_musteri_pivot=z.groupby('Region')['Customer Code'].count().to_frame()
reg_aylik_login=z.groupby('Region')['Ekim Web Regular'].sum().to_frame()
div_musteri_pivot=z.groupby('Division')['Customer Code'].count().to_frame()
div_aylik_login=z.groupby('Division')['Ekim Web Regular'].sum().to_frame()


sum_reg_Heading_list=[
    {"Name":'Müşteri Sayısı',"Pivot":reg_musteri_pivot,"Fillna":0},
    {"Name":'Ekim Login Olan Nokta Sayısı',"Pivot":reg_aylik_login,"Fillna":0}
]


sum_div_Heading_list=[
    {"Name":'Müşteri Sayısı',"Pivot":div_musteri_pivot,"Fillna":0},
    {"Name":'Ekim Login Olan Nokta Sayısı',"Pivot":div_aylik_login,"Fillna":0}
]



sum_reg=Data_leftjoin(z_index_columns,'Region',sum_reg_Heading_list)
sum_reg.reset_index(drop=True, inplace=True)
sum_reg.set_index(keys=['Division'],inplace=True)


l_index_columns= z.pivot_table(index=['Division']).iloc[:,[]].reset_index(drop=False)

sum_div=Data_leftjoin(l_index_columns,'Division',sum_div_Heading_list)
sum_div.reset_index(drop=True, inplace=True)
sum_div.set_index(keys=['Division'],inplace=True)

print("Completed5 MOBIL Sheet Creation:","\t",round(timer()-start,1))


k=z.query('`Web Login`=="Evet"')
web_login_musteri_pivot=k.groupby('Region')['Customer Code'].count().to_frame()

y=z.query('`Mobil Login`=="Evet"')
mobil_login_musteri_pivot=y.groupby('Region')['Customer Code'].count().to_frame()
mobil_login_musteri_pivot

Mobil_Heading_list=[
    {"Name":'Web Login Müşteri Sayısı',"Pivot":web_login_musteri_pivot,"Fillna":0},
    {"Name":'Mobilden Giriş Yapan Müşteri Sayısı',"Pivot":mobil_login_musteri_pivot,"Fillna":0}
]

MOBIL=Data_leftjoin(z_index_columns,'Region',Mobil_Heading_list)

MOBIL['oran'] = MOBIL['Mobilden Giriş Yapan Müşteri Sayısı']/MOBIL['Web Login Müşteri Sayısı']

MOBIL.sort_values(['Division','oran'], ascending=[True, False],inplace=True)
MOBIL.drop('oran',axis =1,inplace=True)

def summarize_mobil(df):
    for k, g in df.groupby('Division', sort=False):
        yield g.append({'Division': str(k)+' Total',
                        'Region': '',
                        'Web Login Müşteri Sayısı': g['Web Login Müşteri Sayısı'].sum(),
                        'Mobilden Giriş Yapan Müşteri Sayısı': g['Mobilden Giriş Yapan Müşteri Sayısı'].sum()}, ignore_index=True)



MOBIL=pd.concat(summarize_mobil(MOBIL), ignore_index=True)
MOBIL.set_index('Division',inplace=True)
MOBIL=MOBIL.assign(Oran=MOBIL['Mobilden Giriş Yapan Müşteri Sayısı']/MOBIL['Web Login Müşteri Sayısı'])

print("Completed5 WEB DUZENLI GIRIS Sheet Creation:","\t",round(timer()-start,1))

_=z.query('`Web Login`=="Evet"')
web_login_musteri_pivot=_.groupby('Region')['Customer Code'].count().to_frame()

_=z.query('`Ekim Web Regular`==1')
duzenli_web_login_musteri_pivot=_.groupby('Region')['Customer Code'].count().to_frame()

_=z.query('`Ekim Web Regular`==1 & `KD Cluster`=="ALTIN"')
duzenli_web_login_Altın_musteri_pivot=_.groupby('Region')['Customer Code'].count().to_frame()

_=z.query('`Ekim Web Regular`==1 & `KD Cluster`=="GÜMÜS"')
duzenli_web_login_Gumus_musteri_pivot=_.groupby('Region')['Customer Code'].count().to_frame()

_=z.query('`Ekim Web Regular`==1 & `KD Cluster`=="PLATIN"')
duzenli_web_login_Platin_musteri_pivot=_.groupby('Region')['Customer Code'].count().to_frame()

duzenli_web_login_musteri_orani_pivot=(duzenli_web_login_musteri_pivot/web_login_musteri_pivot)


web_duzenli_giris_Heading_list=[
    {"Name":'Total Web login',"Pivot":web_login_musteri_pivot,"Fillna":0},
    {"Name":'Düzenli Web Müşteri Sayısı',"Pivot":duzenli_web_login_musteri_pivot,"Fillna":0},
    {"Name":'PLATİN',"Pivot":duzenli_web_login_Platin_musteri_pivot,"Fillna":0},
    {"Name":'ALTIN',"Pivot":duzenli_web_login_Altın_musteri_pivot,"Fillna":0},
    {"Name":'GÜMÜS',"Pivot":duzenli_web_login_Gumus_musteri_pivot,"Fillna":0}

]

WEB_DUZENLİ_GİRİS=Data_leftjoin(z_index_columns,'Region',web_duzenli_giris_Heading_list)

WEB_DUZENLİ_GİRİS['oran'] = WEB_DUZENLİ_GİRİS['Düzenli Web Müşteri Sayısı']/WEB_DUZENLİ_GİRİS['Total Web login']
WEB_DUZENLİ_GİRİS.sort_values(['Division','oran'], ascending=[True, False],inplace=True)
WEB_DUZENLİ_GİRİS.drop('oran',axis =1,inplace=True)


def summarize_web_duzenli(df):
    for k, g in df.groupby('Division', sort=False):
        yield g.append({'Division': str(k)+' Total',
                        'Region': '',
                        'Total Web login': g['Total Web login'].sum(),
                        'Düzenli Web Müşteri Sayısı': g['Düzenli Web Müşteri Sayısı'].sum(),
                        'PLATİN': g['PLATİN'].sum(),
                        'ALTIN': g['ALTIN'].sum(),
                        'GÜMÜS': g['GÜMÜS'].sum()

                       }, ignore_index=True)

WEB_DUZENLİ_GİRİS=pd.concat(summarize_web_duzenli(WEB_DUZENLİ_GİRİS), ignore_index=True)

WEB_DUZENLİ_GİRİS.set_index('Division',inplace=True)
WEB_DUZENLİ_GİRİS=WEB_DUZENLİ_GİRİS.assign(Oran=WEB_DUZENLİ_GİRİS['Düzenli Web Müşteri Sayısı']/WEB_DUZENLİ_GİRİS['Total Web login'])

print("Çıktı Alınıyor...:","\t",round(timer()-start,1))

def dfs_tabs(df_list, sheet_list, file_name):
    writer = pd.ExcelWriter(file_name,engine='xlsxwriter')
    for dataframe, sheet in zip(df_list, sheet_list):
        dataframe.to_excel(writer, sheet_name=sheet, startrow=0 , startcol=0)
        print("Completed",sheet,"\t",round(timer()-start,1))

    writer.save()

# list of dataframes and sheet names
dfs = [sum_reg,sum_div,MOBIL,WEB_DUZENLİ_GİRİS,z]
sheets = ['SumReg','SumDiv','Mobil','Web Düzenli Giriş','Data']



file_name="weblogin_multitest "+str(now)+".xlsx"

dfs_tabs(dfs, sheets, file_name)

print("Completed7-Final",round((timer()-start)/60,1))
print('\a')
