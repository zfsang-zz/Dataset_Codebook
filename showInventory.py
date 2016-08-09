# -*- coding: utf-8 -*-
"""
Created on Mon Jul 11 10:15:03 2016

@author: jsang
"""

import pandas as pd
import glob
import os 
import xlsxwriter
import time

class Inventory(object):
    def reader(self,input,delimiter=None):
        # write in the following formats: xls,xlsx, csv, txt,
        if input.endswith(('.xls','.xlsx','.XLS','.XLSX')):
            tmpDict = pd.read_excel(input,sheetname=None)
            dropList=[]
            for sheet in tmpDict:
                if tmpDict[sheet].empty:
                    dropList.append(sheet)
            for key in dropList:
                del tmpDict[key]
            return tmpDict.values()

        elif input.endswith('.csv'):
            tmp = pd.read_csv(input)
            return [tmp]

        elif input.endswith('.sas7bdat'):
            tmp = pd.read_sas(input)
            return [tmp]
        elif input.endswith(".jmp"):
            return False

    def getFiles(self,pathList,suffix=["xls","xlsx","sas","sas7bdat","csv","jmp"]):
        fileList=[]
        for path in pathList:
            for suffix_single in suffix:
                fileList += glob.glob(path +"/*."+suffix_single)
        return fileList
        
    def showInventory(self,fileList,suffix,name="",droplist=[],outputpath="../output/"):

        df = pd.DataFrame(columns=('Dataset','Folder','Sheet','No.Variables','No.Observation',"ID",'Unique ID','Long Format','Description','Variables','Format','Path','Time','Comment'))
        cnt = 0
        for f in fileList:
            print os.path.basename(f)
            if os.path.basename(f) not in droplist:
                tmpList = self.reader(f) 
                sheet = 1
                if tmpList is False:
                    df.loc[cnt]=[os.path.splitext(os.path.basename(f))[0],'',sheet,"",'','','','','','',os.path.splitext(f)[1],f,time.strftime('%m/%d/%Y %H:%M', time.localtime(os.path.getmtime(f))),'']
                    cnt+=1
                else:
                    for tmpListItem in tmpList:
                        tmpListItem=tmpListItem[pd.notnull(tmpListItem[tmpListItem.columns[0]])]
                        use_cols = [col for col in tmpListItem.columns if 'Unnamed:' not in col]
                        tmpListItem=tmpListItem[use_cols]
                        tmpListItem.dropna(how='all')
                        longformat = 1 if sum(tmpListItem.duplicated(subset = str(list(tmpListItem.columns.values)[0])))>0 else 0
                        df.loc[cnt]=[os.path.splitext(os.path.basename(f))[0],os.path.basename(os.path.dirname(f)),sheet,tmpListItem.shape[1],tmpListItem.shape[0],tmpListItem.columns.values[0],len(tmpListItem[tmpListItem.columns[0]].unique()),longformat,"",','.join(list(tmpListItem.columns)).encode('ascii','ignore'),os.path.splitext(f)[1],f,time.strftime('%m/%d/%Y %H:%M', time.localtime(os.path.getmtime(f))),'']
                        cnt+=1
                        sheet +=1
        if not os.path.exists(outputpath):
            os.makedirs(outputpath)
            
        workbook = xlsxwriter.Workbook(outputpath + name +  "inventory.xlsx")
        worksheet = workbook.add_worksheet()
        # Add the standard url link format.
        url_format = workbook.add_format({
            'font_color': 'blue',
            'underline':  1
        })
        row,col=0,0        
        worksheet.write_row(row,col,list(df.columns))
        row+=1
        for i in range(0,df.shape[0]):
            worksheet.write_url(row,0,df.iloc[i,-3],url_format,df.iloc[i,0],df.iloc[i,-3]); col+=1
            worksheet.write_row(row,1,list(df.ix[i])[1:])
            row+=1
        workbook.close()
#        df.to_csv(outputpath + name +  "inventory.csv",index=False)
#        print outputpath + name +  "inventory.csv"
        
        
    def allFiles(self,path):
        fileList=glob.glob(path +"/*")
        for f in fileList:
            print os.path.basename(f)
        for f in fileList:
            print f   
            
        return fileList
        
    def find(inputPath,output= None,option=1,search=None):
      suffix=(".xls",".xlsx",".sas",".sas7bdat",".csv",".jmp")
      for root, dirs, files in os.walk(inputPath):
          res=[]
          
          for file in files:
              if file.lower().endswith(suffix) and search in file.lower():
                  res+= [os.path.join(root, file)]
      return res


solution=Inventory()
suffix=["xls","xlsx","sas","sas7bdat","csv","jmp"]

solution.showInventory(solution.getFiles("S:\DataCopy"),suffix,"Study_")


