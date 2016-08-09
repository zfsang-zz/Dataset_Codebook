# -*- coding: utf-8 -*-
"""
Created on Fri Jul 15 11:58:57 2016

@author: jsang
"""

import csv
import collections
import pandas as pd
import glob
from dateutil.parser import parse
import os
from datetime import datetime
import xlsxwriter


class VarCodeBook(object):
    
    def type(self,var):
        if pd.isnull(var):
            return "missing"
        
        try: 
            float(str(var))
            return "Number"
        except ValueError:
            pass
        
        try: 
            parse(str(var))
            return "Date"
        except (ValueError, OverflowError) as e:
            pass
        
        return "String"
    
    def Most_Common(self,lst):
        data = collections.Counter(lst)
        return data.most_common(1)[0][0]
    
    def firstValidValue(self,array):
        length = len(array)
        if length<1:
            return float('nan')
        cnt = 0
        while cnt < length:
            if not pd.isnull(array[cnt]):
                return array[cnt]
            else:
                cnt +=1
        return float('nan')
    def getFiles(self,pathList,suffix=["xls","xlsx","sas","sas7bdat","csv","jmp"]):
        fileList=[]
        for path in pathList:
            for suffix_single in suffix:
                fileList += glob.glob(path +"/*."+suffix_single)
        return fileList
        
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
            
    
    def varCodeBookShowAll(self,pathList,suffix,name="",droplist=[],outputpath="../output/"):
        fileList = []
        for path in pathList:
            for suffix_single in suffix:
                fileList += glob.glob(path +"/*."+suffix_single)
            print fileList
        varDict=collections.defaultdict(list)
        for f in fileList:
            print os.path.basename(f)
            if f not in droplist:
                tmpList = self.reader(f) 
                if not(tmpList is False):
                    for tmpListItem in tmpList:
                        for col in tmpListItem.columns.values:
                            if "Unnamed" not in col:
                                if col not in varDict:
                                    varDict[col]= list(tmpListItem[col].unique())
                                else:
                                    varDict[col] = list(set(list(tmpListItem[col].unique()) + varDict[col]))
        for col in varDict:
            tmpUnique=varDict[col]
            length=len(tmpUnique)
            tmpUnique = [i for i in tmpUnique if not pd.isnull(i)]
            missing = 1 if length> len(tmpUnique) else 0
            varType = self.type(self.firstValidValue(tmpUnique))
            varDict[col]=[varType,len(tmpUnique),",".join([str(i) for i in tmpUnique[:5]]),missing]    
            
        writer = csv.writer(open('dict_out.csv', 'wb'))
#        writer.writerow(["Variable Name","Type","Number of Unique Values", "First Five Unique Values","Contains Missing"])

        for key, value in varDict.items():
           writer.writerow([key]+ [val for val in value])
#        if not os.path.exists(outputpath):
#            os.makedirs(outputpath)
#        df.to_csv(outputpath + name +  "inventory.csv",index=False)
        return varDict
        
    def varCodeBookShow(self,fileList,title,primary_key="patient_id",out="codebook.xlsx",outputpath="../output/"):
        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook(out)
        worksheet = workbook.add_worksheet()
        
        # Add a bold format to use to highlight cells.
        bold = workbook.add_format({'bold': 1})
        
        # Add a number format for cells with money.
        money_format = workbook.add_format({'num_format': '$#,##0'})
         
        # Adjust the column width.
        worksheet.set_column(0, 0, 30)
        worksheet.set_column(0, 1, 30)        
        worksheet.set_column(0, 2, 30)        
        mergeBlue_format = workbook.add_format({
            'bold':     True,
            'align':    'center',
            'valign':   'vcenter',
            'fg_color': '#AFC7E7',
        })
        merge_format = workbook.add_format({
            'bold':     False,
            'align':    'center',
            'valign':   'vcenter',
        })
        title_format = workbook.add_format({
            'bold':     True,
            'font_size': 14,
            'align':    'center',
            'valign':   'vcenter',
        })
        field_format = workbook.add_format({
            'bold':     False,
            'align':    'center',
            'valign':   'vcenter',
            'fg_color': '#CFE7B7',
        })
        useful_format = workbook.add_format({
            'bold':     True,
            'align':    'center',
            'valign':   'vcenter',
        })
        url_format = workbook.add_format({
            'font_color': 'blue',
            'underline':  1,
            'valign':   'vcenter',
            'valign':   'center',
        })
        path_format=workbook.add_format({
            'font_size': 13,
            'underline':  0,
            'valign':   'vcenter',
            'valign':   'center',
            'fg_color': '#AFC7E7',

        })
        red=workbook.add_format({
        'font_color':'red',
        })

        # Start from the first cell below the headers.     
        row,col =0, 0
        worksheet.merge_range(row,0,row,2,title,title_format);row+=1
        worksheet.merge_range(row,0,row,2,'',merge_format)
        worksheet.write_url(row,col,"\\".join(fileList[0].split('\\')[:-1]),url_format) ;row+=1
        worksheet.merge_range(row,0,row,2,"Note: highligt cell need more info to provide description, red font color means primary key",merge_format);row+=1
        for f in fileList:        
            varDict=collections.OrderedDict()
            print os.path.basename(f)
            tmpList = self.reader(f)
            if not(tmpList is False):
                for tmpListItem in tmpList:
                    for key in tmpListItem.columns.values:
                        if "Unnamed" not in key:
                            varDict[key]= list(tmpListItem[key].unique())
    
            for key in varDict:
                tmpUnique=varDict[key]
                length=len(tmpUnique)
                tmpUnique = [i for i in tmpUnique if not pd.isnull(i)]
                missing = 1 if length> len(tmpUnique) else 0
                varType = self.type(self.firstValidValue(tmpUnique))
                varDict[key]=[varType,len(tmpUnique),",".join([str(i) for i in tmpUnique[:5]]),missing]
                
            worksheet.write_row(row,col,["","",""]);                                                      row+=1
            worksheet.write_row(row,col,["Useful","","[Comment]"],useful_format);                         row+=1 
            worksheet.merge_range(row,0,row,2,"",mergeBlue_format)
            worksheet.write_url(row,col,f,path_format,os.path.splitext(os.path.basename(f))[0],f); row+=1
            worksheet.write_row(row,col,["Field Name","Data Type","Description"],field_format);           row+=1
             
            for key, val in varDict.items():
                if key == primary_key:
                    worksheet.write_string(row,col,key,red)
                    worksheet.write_string(row,col+1,val[0])
                    row+=1
                else:
                    worksheet.write_row(row,col,[key,val[0],""]); row +=1
            
                
             


lookupDict={"user_id":"user id"}
solution=VarCodeBook()
suffix=["xls","xlsx","sas","sas7bdat","csv","jmp"]
solution.varCodeBookShow(solution.getFiles(["S:\DataCopy"]),"StudyData",primary_key="study_id",out="StudyDataCodeBook.xlsx")