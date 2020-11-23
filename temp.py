# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import os
#import numpy as np
import xlrd
#from datetime import datetime
from datetime import timedelta
#import datetime
#get the name of each sheet/ the number of columns/the number of rows
def getdata(path):
    workbook=xlrd.open_workbook(path)#open workbook
    sheet_name=workbook.sheet_names()
    sheet_row=[]
    sheet_col=[]
    for name in sheet_name:       
      sheet=workbook.sheet_by_name(name)
      sheet_row.append(sheet.nrows)
      sheet_col.append(sheet.ncols)
   # print(sheet_name,sheet_row,sheet_col)
    return (sheet_name,sheet_row,sheet_col)

# get data summary based on month    
def getMonthcontent(path, sheet_name,revised_day):
    sheet_content=pd.read_excel(path, sheet_name=sheet_name, header=None)
    sheet_content[0]=sheet_content[0].astype('datetime64[ns]')
    delta=timedelta(days=revised_day)
    sheet_content[0]=sheet_content[0]-delta
    sheet_content['Month'] = sheet_content[0].apply(lambda x:x.month)
    sheet_content_Month = sheet_content.groupby(['Month'])[[2]].agg('sum')     
    sheet_content_Month = sheet_content_Month.T    
    return (sheet_content_Month)

#get data summary based on Week
def getWeekcontent(path, sheet_name,revised_day):
    sheet_content=pd.read_excel(path, sheet_name=sheet_name, header=None)
    sheet_content[0]=sheet_content[0].astype('datetime64[ns]')
    delta=timedelta(days=revised_day)
    sheet_content[0]=sheet_content[0]-delta
    sheet_content['Week'] = sheet_content[0].apply(lambda x:x.week)
    sheet_content_Week = sheet_content.groupby(['Week'])[[2]].agg('sum')   
    sheet_content_Week=sheet_content_Week.T
    #s.t week numbers become continuous
    sheet_week_index=sheet_content_Week.columns.values.tolist()
    new_index=[]
    new_index.append(sheet_week_index[0])
    for i in range(len(sheet_week_index)-2):
        if sheet_week_index[i+1]==sheet_week_index[i]+1:
           new_index.append(sheet_week_index[i+1])
        else:
            new_index.append(sheet_week_index[i]+1)
            new_index.append(sheet_week_index[i+1])
    new_sheet_content_Week=pd.DataFrame(columns=new_index)
    for index in sheet_week_index:
       new_sheet_content_Week[index]=sheet_content_Week[index]   
    #return (sheet_content_Week)
    return (new_sheet_content_Week)

# calculate rate
def rate(datalist):
    rate=[]
    rate.append(" ")
    for i in range(len(datalist)-1):
        if datalist[i]==0:
            data=" "
        else:
            data=(datalist[i+1]-datalist[i])/datalist[i]
            data="%.2f%%"%(data*100)
        rate.append(data)
    return(rate)

if __name__=='__main__':
  for i in range(0,10):
        Inputpath=input("Please input file path:")
        if Inputpath=="END":
            print("End Process")
            break
        sheet_name=getdata(Inputpath)[0]
        #column_name=["July","Aug","Sep","Oct","Nov","Dec"]
        revised_day=int(input("input days need to be subtracted:"))
        total_content_Month=pd.DataFrame(columns=sheet_name)
        total_content_Week=pd.DataFrame(columns=sheet_name)

        for name in sheet_name:
            sheet_content_Month=getMonthcontent(Inputpath,name,revised_day)
            total_content_Month[name]=sheet_content_Month.iloc[0]
            sheet_content_Week=getWeekcontent(Inputpath,name,revised_day)
            total_content_Week[name]=sheet_content_Week.iloc[0]



        total_content_Month=total_content_Month.T
        nan_total_content_Month = total_content_Month.fillna(0)
        total_content_Month['Total'] = nan_total_content_Month.apply(lambda x: x.sum(), axis=1)
        total_content_Month['Rate']=rate(total_content_Month.iloc[:,-1].tolist())

        total_content_Week=total_content_Week.T
        nan_total_content_Week=total_content_Week.fillna(0)
        total_content_Week['Total']=nan_total_content_Week.apply(lambda x:x.sum(),axis=1)
        total_content_Week['Rate']=rate(total_content_Week.iloc[:,-1].tolist())

     #write results into new file
        outputpath=os.path.dirname(Inputpath)+'\\'
        outputfile=input('Please input file name which results wrote into:')
        Outputpath=outputpath+outputfile+'.xlsx'
        writer = pd.ExcelWriter(Outputpath)
        total_content_Month.to_excel(excel_writer=writer,sheet_name='RBCW_Month',index=True)
        total_content_Week.to_excel(excel_writer=writer,sheet_name='RBCW_Week',index=True)
        writer.save()
        writer.close()

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    