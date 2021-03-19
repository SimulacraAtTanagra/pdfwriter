# -*- coding: utf-8 -*-
"""
Created on Thu May 28 13:11:39 2020

@author: sayers
"""
import os
import pdfrw
import pandas as pd
import subprocess
import pyautogui as pig
from time import sleep
from emailautosend import mailthat

from datetime import date 
today=date.today().strftime("%B %d, %Y")

def colclean(df):
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
    return(df)

#INVOICE_TEMPLATE_PATH = 's:\\downloads\\20_21_classified-hourly_reappointment.pdf'
#INVOICE_OUTPUT_PATH = 'testpdf2.pdf'

ANNOT_KEY = '/Annots'
ANNOT_FIELD_KEY = '/T'
ANNOT_VAL_KEY = '/V'
ANNOT_RECT_KEY = '/Rect'
SUBTYPE_KEY = '/Subtype'
WIDGET_SUBTYPE_KEY = '/Widget'

def write_fillable_pdf(input_pdf_path, output_pdf_path, data_dict):
    ANNOT_KEY = '/Annots'
    ANNOT_FIELD_KEY = '/T'
    ANNOT_VAL_KEY = '/V'
    ANNOT_RECT_KEY = '/Rect'
    SUBTYPE_KEY = '/Subtype'
    WIDGET_SUBTYPE_KEY = '/Widget'
    template_pdf = pdfrw.PdfReader(input_pdf_path)
    annotations = template_pdf.pages[0][ANNOT_KEY]
    for annotation in annotations:
        if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
            if annotation[ANNOT_FIELD_KEY]:
                key = annotation[ANNOT_FIELD_KEY][1:-1]
                if key in data_dict.keys():
                    try:
                        annotation.update(
                        pdfrw.PdfDict(V='{}'.format(data_dict[key]))
                    )
                        annotation.update(pdfrw.PdfDict(Ff=1))
                    except:
                        print("didn't work boss")

    template_pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true'))) 
    pdfrw.PdfWriter().write(output_pdf_path, template_pdf)
 #df=df.reset_index()
#with df from cjrplatform

emails=pd.read_excel('s://downloads//addys2.xls')
'''emails2=pd.read_excel('s://downloads//HR_REPORTS_EMAILS_ALL_2831.xlsx')
emails2=colclean(emails2)
emailemps=emails2[['id']]
emailemps=emailemps.drop_duplicates()
emailemps=colclean(emailemps)
def getother(empl):
    xyz=list(emails2[emails2.id==empl].email.unique())
    try:
        return([x for x in xyz if 'york.cuny.edu' in x][0])
    except:
        pass
    

emailemps[['ycemails']]=emailemps.id.apply(getother)
emailemps[['emails']]=emailemps.id.apply(getother,args=(,'no'))'''

employees= pd.read_excel('s://downloads//resp2.xls')
distlist=pd.read_excel('s://downloads//distlist.xls')
employees.columns=employees.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
emails.columns = emails.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
distlist.columns = distlist.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
df = employees.merge(emails,how="left",left_on="empl_id",right_on="id")
ddf = df.merge(distlist,how="left")
ddf12="yes"
ddf.columns
#ddf2=pd.read_excel('s://downloads//CA_rt.xlsx')
#ddf2.columns = ddf2.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
#dist=ddf2[["dept_id_job","area"]].drop_duplicates()
data_dict={}
counter=0
for i in ddf.index.values:
    if counter<1000:
        
        try:
            data_dict['emplid']= f'{ddf.iloc[i].empl_id} '
        except:
            print(f"emplid didn't load into dict. {ddf.iloc[i]}")
        try:    
            data_dict['name']= f'{ddf.iloc[i].person_nm}: '
        except:
            pass
        try:
            data_dict['title']= f'{ddf.iloc[i].labor_job_ld} '
        except:
            pass
        try:
            data_dict['rate']= f'${ddf.iloc[i].comp_rt}/hr '
        except:
            pass
        if len(f'{ddf.iloc[i].comp_rt}')==4:
            try:
                data_dict['rate']= f'${ddf.iloc[i].comp_rt}0/hr '
            except:
                pass
        
        #'division': f'{df.iloc[i].division}',
        try:
            data_dict['department']= f'{ddf.iloc[i].dept_descr_job} '
        except:
            pass
        if len(f'{ddf.iloc[i].busn}')<4:
            try:
                data_dict['email']=""
            except:
                pass
        else:
            data_dict['email']= f"{ddf.iloc[i].busn} "
        try:
            data_dict['division']= f'{ddf.iloc[i].area} '
        except:
            pass
        try:
            data_dict['rt']= f"{ddf.iloc[i].sup_nam} "
        except:
            pass
#        try:
#           data_dict['letterdatee']= f'{today} '
#        except:
#            pass
        INVOICE_TEMPLATE_PATH = 's:\\downloads\\20_21_classified-hourly_reappointment.pdf'
        INVOICE_OUTPUT_PATH = f's:\\desktop\\September_Letters\\{ddf.iloc[i].empl_id}_{ddf.iloc[i].dept_id_job}_Sept_letter.pdf'
    
        try:
            write_fillable_pdf(INVOICE_TEMPLATE_PATH, INVOICE_OUTPUT_PATH, data_dict)
        except:
            print("Didn't work boss. Don't know why.")
        counter+=1
countthis=0
ddf.camp=ddf.camp.astype('str')
ddf.busn=ddf.busn.astype('str')
ddf.othr=ddf.othr.astype('str')
ddf.home=ddf.home.astype('str')
ddf.dorm=ddf.dorm.astype('str')

subject="September Reappointment Letter"
flag=1
misslist=[]
for i in ddf.index.values:
    try:
        obj = f's:\\desktop\\September_Letters\\{ddf.iloc[i].empl_id}_{ddf.iloc[i].dept_id_job}_Sept_letter.pdf'
    except:
        print(f"{ddf.iloc[i].person_nm}'s letter is not here")
        pass
    if ddf.iloc[i].labor_job_ld in ['EOC Assistant','College Assistant']:
        cc = f'{ddf.iloc[i].rtemail};ajackson1@york.cuny.edu'
    else:
        cc = f'{ddf.iloc[i].rtemail};mwilliams@york.cuny.edu'

    if len(ddf.iloc[i].busn)>4:
        to=ddf.iloc[i].busn
    elif len(ddf.iloc[i].camp)>4:
        to=ddf.iloc[i].camp
    elif len(ddf.iloc[i].othr)>4:
        to=ddf.iloc[i].othr
    elif len(ddf.iloc[i].home)>4:
        to=ddf.iloc[i].home
    else:
        to=ddf.iloc[i].dorm
    bcc= f'{df.iloc[i].camp};{df.iloc[i].othr};{df.iloc[i].home};{df.iloc[i].dorm}'
    bcc=bcc.replace("nan","")
    bcc=''.join(bcc)
    if len(bcc)<1:
        bcc=""
    if len(str(ddf.iloc[i].busn))<5 and len(str(ddf.iloc[i].camp))<5 and len(str(ddf.iloc[i].othr))<5 and len(str(ddf.iloc[i].home))<5:
        print(f"{ddf.iloc[i].person_nm}'s letter exists but could not be sent due to not having any accounts")
        misslist.append(ddf.iloc[i].person_nm)
        continue
    if flag==1:
        mailthat(to,cc,bcc,obj,subject)
        print(f'{ddf.iloc[i].empl_id} sent!')
    else:
        print(f'{to},{cc},{bcc},{subject}')
        countthis+=1
        print(countthis)
print(misslist)
#deprecating this bit asI can do it in above segment now
"""
directory_in_str = 's:\\desktop\\testpdfs'
directory = os.fsencode(directory_in_str)           #defines directory as indicated string
os.chdir(directory)                                 #navigate to directory specified
for file in os.listdir(directory):                  #iterates over all the files here
    filename = os.fsdecode(file)
    subprocess.Popen([filename],shell=True)
    sleep(2)
    pig.click(454,447)
    for i in range(8):
        pig.press('end')
        pig.typewrite(' ')
        pig.press('tab')
    
    pig.hotkey('ctrl','s')
    sleep(.3)
    pig.hotkey('ctrl','f4')
sleep(3)
pig.position()
"""