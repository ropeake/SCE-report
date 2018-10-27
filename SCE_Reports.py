# -*- coding: utf-8 -*-
"""
Created on Sat Mar 17 16:39:32 2018

@author: Ro
"""

import pandas as pd
import win32ui
import win32con
o=win32ui.CreateFileDialog(1)
if o.DoModal()==1:
    filename=o.GetPathName()
#%% Creating the Programme Name Data Frame

    
#%% Import CSV file
    df=pd.read_csv(filename, encoding='latin1', sep='|')
  #                 names=['SCJ_Code','Forenames','Surname','Gender','Course_Title','Course_Code','Route_Code','SCE_Year',
  #                        'Status','MOA','Block','Dept_Code','Dept','Fac_Code','Fac','Home_Email','KCL_Email','Username',
  #                        'HESA_Start_Date','Exp_End_Date','SPR_Batch','SCE_Batch','PRS_1','PRS_Name','PRS_2','PRS_2_Name'])
 #   df.transaction_date=df.description.str.split(' ON ',expand=True)[1] #split and then return column 1 (firts column is 0)

#create a new series (column), use square brackets in case it has a space in it, tell it which serires you're splitting
    df['Student_Number']=df['SCJ Code'].str.split('/',expand=True)[0] #split (tell it what to split on) and then return column [0] (firts column)
#create a dictionary to look up the programme Level
#    df['Level']=
    df['Course&Route']=df['Course Code']+df['Route Code']
#create a dictionary to look up the programme name
    CourseRouteName = {
            'UGDP1CSPHSPH': 'GRAD DIP',#strings must be indicated with'' not necessary for numbers
            'UBSH3CJMMPHJMMPH': 'MATH PHY BSC',
            'UMSH4CJMMPHJMMPH': 'MATH PHY MSCI',
            'TMSC1CTPHYTPHY': 'MSC',
            'TMSC2CTPHYTPHY': 'MSC - PT',
            'UBSH3CMPHMAMPHMA': 'PHY MED APP BSC',
            'UBSH4CJPHLYJPHLY': 'PHY PHIL AB BSC',
            'UBSH3CJPHPLJPHPL': 'PHY PHIL BSC',
            'UBSH3CMPHTPMPHTP': 'PHY THE PHY BSC',
            'UMSH4CMPHTPMPHTP': 'PHY THE PHY MSCI',
            'UBSH4CSPDSPD': 'PHY YR AB BSC',
            'UBSH3CSPHSPH': 'PHYSICS BSC',
            'UMSH4CSPHSPH': 'PHYSICS MSCI',
            'TCDT1CTPHYTPHY': 'Physics Non-Award',
            'UMSH4CJPHPIJPHPI': 'PHY PHIL MSCI',
            'RNCR1CRPHY': 'PHYSICS NONCRD',
            'RDPL3CRPHY': 'PHYSICS M/PHD',
            'RDPL4CRPHY': 'PHYSICS M/PHD',
            'UBSH3CSPHMPHAC': 'PHYSICS BSC (Astro)',
            'UMSH4CSPHMPHAC': 'PHYSICS MSCI (Astro)',
}

    df['CourseRouteName'] = df['Course&Route'].map(CourseRouteName)#create a new column, match the course route code field to the course route name in the dictionary return the name into the new column
#   df['Student_Number']=df['SCJ Code'].str.split('/',expand=True)[0]
    df['Year of Start']=df['HESA Start Date'].str.split('/',expand=True)[2]#create a new Series for the year of start
    df=df[['Student_Number','SCJ Code','Forenames','Surname','Gender','Course Title','Course Code','Route Code','SCE Year','Status','MOA','Block',
           'Dept Code','Dept','Fac Code','Fac','Home Email','KCL Email','Username','HESA Start Date','Exp End Date',
           'SPR Batch','SCE Batch','PRS 1','PRS Name','PRS 2',' PRS 2 Name','Course&Route','CourseRouteName','Year of Start']]
#%% Code to remove students who appear on list twice
#    df=df.drop_duplicates(keep='first')
    
#%% Code to copy the list to the clipboard
    df.to_clipboard(excel=True,index=False,header=True,sep='\t')


#%%
else:
    win32ui.MessageBox('You have chosen to exit','SCE Report Formatter',win32con.MB_ICONSTOP)
    
    
    