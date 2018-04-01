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
#%%
    df=pd.read_csv(filename,'|',header=0,
                   names=['SCJ_Code','Forenames','Surname','Gender','Course_Title','Course_Code','Route_Code','SCE_Year','Status','MOA','Block','Dept_Code','Dept','Fac_Code','Fac','Home_Email','KCL_Email','Username','HESA_Start_Date','Exp_End_Date','SPR_Batch','SCE_Batch','PRS_1','PRS_Name','PRS_2','PRS_2_Name'])
#    df.Student_Number=df.SCJ_Code.rstrip('/')    
    df.to_clipboard(excel=True,index=False,header=True)
#%%
else:
    win32ui.MessageBox('You have chosen to exit','SCE Report Formatter',win32con.MB_ICONSTOP)
    
    
    