import os
import sys
print(sys.version)
import csv
import math
import pandas as pd
import numpy as np
import datetime
import xlwt
import glob
import pyexcel as p
#import Impact&VerbatimFunctions
import AdminImpact, EndUserImpact

print(sys.argv)
program_name = sys.argv[0]
report_type = sys.argv[1]
data_file = sys.argv[2]
time_period = sys.argv[3]
period_1 = sys.argv[4]
period_2 = sys.argv[5]

# class Person: #Class Object Template, used for reference
#   def __init__(self, name, age):
#     self.name = name
#     self.age = age

#   def myfunc(self):
#     print("Hello my name is " + self.name)

# p1 = Person("John", 36)
# p1.myfunc()

# data_file = 'EndUserApp(April_May).csv'
# build_file = '../Downloads/Build_to_fork_May2019.csv'

class Impact:
    ''' 
    Docstring for Impact Class Object for generating Impact reports

    INPUTS
    kind: Specify whether you want "Admin" or "EndUser" Impact Analysis
    time_interval: Specify time interval to perform Impact Analysis on (usually Fork, Month, or Week)
        Note -  If you are looking to filter impact by custom period (e.g. every two weeks), you will need to define it in custom field in your dataset. 
    previous_interval: Initial Period for analysis
    current_interval: Next Period for analysis
    '''
    def __init__(self, kind, time_interval,previous_interval,current_interval):
        self.kind = kind #"EndUser" or "Admin"
        self.time = time_interval #Fork, Month, etc.
        self.prev = previous_interval #Initial Period for analysis
        self.curr = current_interval #Next Period for analysis

    def EndUser(self, data_file, build_file, survey="Default (Both)"):
        ''' If you do not Specify "App" as survey input (do this if your data has no O365/Suite data), this report will run for both '''
        self.survey = str(survey) #"Suite" or "App"
        if self.kind == 'EndUser':
            if self.survey=='App': #No Suite Data 
                df = EndUserImpact.create_NPSdataframe(data_file,build_file)
                EndUserImpact.custom_impact_pivot_levels(df,self.time,self.survey,self.prev,self.curr)
            else:
                df = EndUserImpact.create_NPSdataframe(data_file,build_file)
                EndUserImpact.custom_impact_pivot_allup(df,self.time,self.prev,self.curr)
                EndUserImpact.custom_impact_pivot_levels(df,self.time,self.survey,self.prev,self.curr)
            
    def Admin(self,data_file):
        if self.kind == 'Admin':
            admin_df = AdminImpact.create_NPSdataframe(data_file)
            AdminImpact.custom_impact_pivot_admin(admin_df,self.time,self.prev,self.curr)

    def get_timestamp(self):
        return self.time.split('_')[0].upper()



if report_type == 'EndUser':
    fork_file = sys.argv[6]
    s_o_a = sys.argv[7]
    reportEU = Impact("EndUser",time_period,[period_1],[period_2])
    stamp = reportEU.get_timestamp()
    print(stamp)
    reportEU.EndUser(data_file,fork_file,s_o_a)

    d = '('+str(datetime.datetime.today().month)+'-'+str(datetime.datetime.today().day)+'-'+str(datetime.datetime.today().year)+')'
    num_check = set('0123456789.')
    letter_check = set('qwertyuiopasdfghjklzxcvbnm-+()[]{}')

    wb = xlwt.Workbook()
    for filename in glob.glob(str(s_o_a+'*')+stamp+".csv"):
         (f_path, f_name) = os.path.split(filename)
         (f_short_name, f_extension) = os.path.splitext(f_name)
         ws = wb.add_sheet(f_short_name.partition(stamp)[0])
         fileReader = csv.reader(open(filename, 'rt'))
         for rowx, row in enumerate(fileReader):
             for colx, value in enumerate(row):
                 if (any((n in value) for n in num_check)) & ('.0.' not in value) & (any((c in value.lower()) for c in letter_check) is False):
                     #print(value)
                     num_value = float(value)
                     ws.write(rowx, colx, num_value)
                 elif value.strip().startswith('-'):
                     num_value = float(value)
                     ws.write(rowx, colx, num_value)
                 else:
                     ws.write(rowx, colx, value) 

    wb.save(s_o_a+"ImpactReports"+stamp+d+".xls")
    p.save_book_as(file_name=s_o_a+'ImpactReports'+stamp+d+'.xls',
                    dest_file_name=s_o_a+'ImpactReports'+stamp+d+'.xlsx')

if report_type == 'Admin':
    reportA = Impact("Admin",time_period,[period_1],[period_2])
    stamp = reportA.get_timestamp()
    print(stamp)
    reportA.Admin(data_file)







               