import pandas as pd
import datetime as datetime
import numpy as np
import os
#os.getcwd()
import sys
print(sys.version)
import csv
import xlwt
import glob
import pyexcel as p

#Script by Jared Anderson

path = 'ForkAnalysisDump/'

srq_files_separate = True
app_files_separate = True

Build= pd.read_csv('C:/Users/v-Jaand/Downloads/Build_to_fork_June2019.csv')
Builds = Build.set_index('ForkMth')['OfficeForkBuild'].to_dict()
Build =Build.set_index('OB_ThirdPart')['ForkMth'].to_dict()

d = '('+str(datetime.datetime.today().month)+'-'+str(datetime.datetime.today().day)+'-'+str(datetime.datetime.today().year)+')'
num_check = set('0123456789.')
letter_check = set('qwertyuiopasdfghjklzxcvbnm-+()[]{}')

if srq_files_separate==True:
    if app_files_separate == True:
        for srq in ['Suite','App']:
            for app in ['Outlook','Excel','PowerPoint','Word']:
                iw= pd.read_csv(path+srq+app+'ForkMonthData.csv'#,index_col= 'ProcessSessionId'
                    )
                iw.columns = [x.partition('_')[2] for x in iw.columns.to_list()]
                iw['minorbuild'] = iw['OfficeBuild'].str.split('.', expand=True)[2].astype(float)
                iw['fork'] = iw['minorbuild'].astype(float).map(Build)
                # Convert LongDate to a datetime
                iw["LongDateTime"] = iw["DateTime"].apply(lambda x: pd.to_datetime(x))
                iw['Feed_year'] = pd.DatetimeIndex(iw['LongDateTime']).year
                iw['month_year'] = iw.LongDateTime.dt.to_period('M')
                iw['week_year'] = iw.LongDateTime.dt.to_period('W-SAT')

                iw3 = iw[iw['fork']>=iw['fork'].max()-11][iw['month_year']>=iw['month_year'].max()-11]

                # iw3['Feed_month'] = pd.DatetimeIndex(iw3['LongDateTime']).month
                # iw3['Feed_week'] = pd.DatetimeIndex(iw3['LongDateTime']).week
                'Analysis dates: {0} to {1}'.format(iw3["LongDateTime"].min().strftime('%Y-%m-%d'), iw3["LongDateTime"].max().strftime('%Y-%m-%d'))

                # iw3.loc[iw3['SurveyRatingQuestion'] == 'How likely are you to recommend Office 365 to a friend or colleague?', 'SurveyType'] = 'Suite'
                # iw3.loc[iw3['SurveyRatingQuestion'] == 'How likely are you to recommend our application to a friend or colleague?', 'SurveyType'] = 'App'    
                iw3.loc[iw3['SurveyName'].str.startswith('Suite'), 'SurveyType'] = 'Suite'
                iw3.loc[iw3['SurveyName'].str.startswith('App'), 'SurveyType'] = 'App'
                iw3['AppType'] = iw3['SurveyType'] + iw3['App']
                iw3['SurveyApp'] = iw3['SurveyType'] + iw3['App']

                iw3.loc[iw3['Rating'] == 4, 'ratingvalue'] = 0
                iw3.loc[iw3['Rating'] == 5, 'ratingvalue'] = 100
                iw3.loc[iw3['Rating']==3, 'ratingvalue']=-100
                iw3.loc[iw3['Rating']==2, 'ratingvalue']=-100
                iw3.loc[iw3['Rating']==1, 'ratingvalue']=-100

                AppNPS= (iw3.pivot_table(columns=['month_year'], 
                                        index=['fork'], 
                                        aggfunc=[np.mean,len], values='ratingvalue', margins=True)
                    #.add_prefix('NPS_')
                    #.reset_index()
                )
                AppNPSByType= (iw3.pivot_table(columns=['month_year'], 
                                        index=['UserType','fork'], 
                                        aggfunc=[np.mean,len], values='ratingvalue', margins=True)
                    #.add_prefix('NPS_')
                    #.reset_index()
                    )
                AppNPS['mean'] = AppNPS['mean'].round(1)
                AppNPSByType['mean'] = AppNPSByType['mean'].round(1)

                AppNPS.rename(index=Builds, columns={'mean':'NPS','len':'Count'},inplace=True)
                AppNPSByType.rename(index=Builds, columns={'mean':'NPS','len':'Count'},inplace=True)

                AppNPS.to_csv(path+srq+app+'_Fork_Monthly'+d+'.csv')
                pd.DataFrame(data=['BY USER TYPE']).to_csv(app+'_Fork_Monthly'+d+'.csv',mode='a')
                AppNPSByType.to_csv(path+srq+app+'_Fork_Monthly'+d+'.csv',mode='a')

            wb = xlwt.Workbook()
            for filename in glob.glob(path+srq+"*_Fork_Monthly"+d+".csv"):
                (f_path, f_name) = os.path.split(filename)
                (f_short_name, f_extension) = os.path.splitext(f_name)
                ws = wb.add_sheet(f_short_name.partition(d)[0])
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
                            
            wb.save(path+srq+"ForkMonthAnalysis"+d+".xls")
            p.save_book_as(file_name=path+srq+'ForkMonthAnalysis'+d+'.xls',
                            dest_file_name=path+srq+'ForkMonthAnalysis'+d+'.xlsx')
    if app_files_separate == False:
        for srq in ['Suite','App']:
            iw= pd.read_csv(path+srq+'ForkMonthData.csv'#,index_col= 'ProcessSessionId'
                )
            iw.columns = [x.partition('_')[2] for x in iw.columns.to_list()]
            iw['minorbuild'] = iw['OfficeBuild'].str.split('.', expand=True)[2]
            iw['fork'] = iw['minorbuild'].astype(float).map(Build)
            # Convert LongDate to a datetime
            iw["LongDateTime"] = iw["DateTime"].apply(lambda x: pd.to_datetime(x))
            iw['Feed_year'] = pd.DatetimeIndex(iw['LongDateTime']).year
            iw['month_year'] = iw.LongDateTime.dt.to_period('M')
            iw['week_year'] = iw.LongDateTime.dt.to_period('W-SAT')

            iw3 = iw[iw['fork']>=iw['fork'].max()-11][iw['month_year']>=iw['month_year'].max()-11]

            # iw3['Feed_month'] = pd.DatetimeIndex(iw3['LongDateTime']).month
            # iw3['Feed_week'] = pd.DatetimeIndex(iw3['LongDateTime']).week
            'Analysis dates: {0} to {1}'.format(iw3["LongDateTime"].min().strftime('%Y-%m-%d'), iw3["LongDateTime"].max().strftime('%Y-%m-%d'))

            # iw3.loc[iw3['SurveyRatingQuestion'] == 'How likely are you to recommend Office 365 to a friend or colleague?', 'SurveyType'] = 'Suite'
            # iw3.loc[iw3['SurveyRatingQuestion'] == 'How likely are you to recommend our application to a friend or colleague?', 'SurveyType'] = 'App'    
            iw3.loc[iw3['SurveyName'].str.startswith('Suite'), 'SurveyType'] = 'Suite'
            iw3.loc[iw3['SurveyName'].str.startswith('App'), 'SurveyType'] = 'App'
            iw3['AppType'] = iw3['SurveyType'] + iw3['App']
            iw3['SurveyApp'] = iw3['SurveyType'] + iw3['App']

            iw3.loc[iw3['Rating'] == 4, 'ratingvalue'] = 0
            iw3.loc[iw3['Rating'] == 5, 'ratingvalue'] = 100
            iw3.loc[iw3['Rating']==3, 'ratingvalue']=-100
            iw3.loc[iw3['Rating']==2, 'ratingvalue']=-100
            iw3.loc[iw3['Rating']==1, 'ratingvalue']=-100
            if 'SurveyType' in list(iw3.columns):
                AppNPS= (iw3.pivot_table(columns=['month_year'], 
                                        index=['SurveyType','fork'], 
                                        aggfunc=[np.mean,len], values='ratingvalue', margins=True)
                    #.add_prefix('NPS_')
                    #.reset_index()
                )
                AppNPSByType= (iw3.pivot_table(columns=['month_year'], 
                                        index=['SurveyType','UserType','fork'], 
                                        aggfunc=[np.mean,len], values='ratingvalue', margins=True)
                    #.add_prefix('NPS_')
                    #.reset_index()
                    )
                AppNPS['mean'] = AppNPS['mean'].round(1)
                AppNPSByType['mean'] = AppNPSByType['mean'].round(1)

                AppNPS.rename(index=Builds, columns={'mean':'NPS','len':'Count'},inplace=True)
                AppNPSByType.rename(index=Builds, columns={'mean':'NPS','len':'Count'},inplace=True)

                AppNPS.to_csv(path+srq+app+'_Fork_Monthly'+d+'.csv')
                pd.DataFrame(data=['BY USER TYPE']).to_csv(app+'_Fork_Monthly'+d+'.csv',mode='a')
                AppNPSByType.to_csv(path+srq+app+'_Fork_Monthly'+d+'.csv',mode='a')

if srq_files_separate==False:
    if app_files_separate == True:
        for app in ['Outlook','Excel','PowerPoint','Word']:
                iw= pd.read_csv(path+srq+app+'ForkMonthData.csv'#,index_col= 'ProcessSessionId'
                    )
                iw.columns = [x.partition('_')[2] for x in iw.columns.to_list()]
                iw['minorbuild'] = iw['OfficeBuild'].str.split('.', expand=True)[2]
                iw['fork'] = iw['minorbuild'].astype(float).map(Build)
                # Convert LongDate to a datetime
                iw["LongDateTime"] = iw["DateTime"].apply(lambda x: pd.to_datetime(x))
                iw['Feed_year'] = pd.DatetimeIndex(iw['LongDateTime']).year
                iw['month_year'] = iw.LongDateTime.dt.to_period('M')
                iw['week_year'] = iw.LongDateTime.dt.to_period('W-SAT')

                iw3 = iw[iw['fork']>=iw['fork'].max()-11][iw['month_year']>=iw['month_year'].max()-11]

                # iw3['Feed_month'] = pd.DatetimeIndex(iw3['LongDateTime']).month
                # iw3['Feed_week'] = pd.DatetimeIndex(iw3['LongDateTime']).week
                'Analysis dates: {0} to {1}'.format(iw3["LongDateTime"].min().strftime('%Y-%m-%d'), iw3["LongDateTime"].max().strftime('%Y-%m-%d'))

                # iw3.loc[iw3['SurveyRatingQuestion'] == 'How likely are you to recommend Office 365 to a friend or colleague?', 'SurveyType'] = 'Suite'
                # iw3.loc[iw3['SurveyRatingQuestion'] == 'How likely are you to recommend our application to a friend or colleague?', 'SurveyType'] = 'App'    
                iw3.loc[iw3['SurveyName'].str.startswith('Suite'), 'SurveyType'] = 'Suite'
                iw3.loc[iw3['SurveyName'].str.startswith('App'), 'SurveyType'] = 'App'
                iw3['AppType'] = iw3['SurveyType'] + iw3['App']
                iw3['SurveyApp'] = iw3['SurveyType'] + iw3['App']

                iw3.loc[iw3['Rating'] == 4, 'ratingvalue'] = 0
                iw3.loc[iw3['Rating'] == 5, 'ratingvalue'] = 100
                iw3.loc[iw3['Rating']==3, 'ratingvalue']=-100
                iw3.loc[iw3['Rating']==2, 'ratingvalue']=-100
                iw3.loc[iw3['Rating']==1, 'ratingvalue']=-100

                AppNPS= (iw3.pivot_table(columns=['month_year'], 
                                        index=['fork'], 
                                        aggfunc=[np.mean,len], values='ratingvalue', margins=True)
                    #.add_prefix('NPS_')
                    #.reset_index()
                )
                AppNPSByType= (iw3.pivot_table(columns=['month_year'], 
                                        index=['UserType','fork'], 
                                        aggfunc=[np.mean,len], values='ratingvalue', margins=True)
                    #.add_prefix('NPS_')
                    #.reset_index()
                    )
                AppNPS['mean'] = AppNPS['mean'].round(1)
                AppNPSByType['mean'] = AppNPSByType['mean'].round(1)

                AppNPS.rename(index=Builds, columns={'mean':'NPS','len':'Count'},inplace=True)
                AppNPSByType.rename(index=Builds, columns={'mean':'NPS','len':'Count'},inplace=True)

                AppNPS.to_csv(path+srq+app+'_Fork_Monthly'+d+'.csv')
                pd.DataFrame(data=['BY USER TYPE']).to_csv(app+'_Fork_Monthly'+d+'.csv',mode='a')
                AppNPSByType.to_csv(path+srq+app+'_Fork_Monthly'+d+'.csv',mode='a')


    if app_files_separate == False:
        iw= pd.read_csv('ForkMonthData.csv'#,index_col= 'ProcessSessionId'
                )
        iw.columns = [x.partition('_')[2] for x in iw.columns.to_list()]
        iw['minorbuild'] = iw['OfficeBuild'].str.split('.', expand=True)[2]
        iw['fork'] = iw['minorbuild'].astype(float).map(Build)
        # Convert LongDate to a datetime
        iw["LongDateTime"] = iw["DateTime"].apply(lambda x: pd.to_datetime(x))
        iw['Feed_year'] = pd.DatetimeIndex(iw['LongDateTime']).year
        iw['month_year'] = iw.LongDateTime.dt.to_period('M')
        iw['week_year'] = iw.LongDateTime.dt.to_period('W-SAT')

        iw3 = iw[iw['fork']>=iw['fork'].max()-11][iw['month_year']>=iw['month_year'].max()-11]

        # iw3['Feed_month'] = pd.DatetimeIndex(iw3['LongDateTime']).month
        # iw3['Feed_week'] = pd.DatetimeIndex(iw3['LongDateTime']).week
        'Analysis dates: {0} to {1}'.format(iw3["LongDateTime"].min().strftime('%Y-%m-%d'), iw3["LongDateTime"].max().strftime('%Y-%m-%d'))

        # iw3.loc[iw3['SurveyRatingQuestion'] == 'How likely are you to recommend Office 365 to a friend or colleague?', 'SurveyType'] = 'Suite'
        # iw3.loc[iw3['SurveyRatingQuestion'] == 'How likely are you to recommend our application to a friend or colleague?', 'SurveyType'] = 'App'    
        iw3.loc[iw3['SurveyName'].str.startswith('Suite'), 'SurveyType'] = 'Suite'
        iw3.loc[iw3['SurveyName'].str.startswith('App'), 'SurveyType'] = 'App'
        iw3['SurveyApp'] = iw3['SurveyType'] + iw3['App']

        iw3.loc[iw3['Rating'] == 4, 'ratingvalue'] = 0
        iw3.loc[iw3['Rating'] == 5, 'ratingvalue'] = 100
        iw3.loc[iw3['Rating']==3, 'ratingvalue']=-100
        iw3.loc[iw3['Rating']==2, 'ratingvalue']=-100
        iw3.loc[iw3['Rating']==1, 'ratingvalue']=-100
        if 'AppType' in list(iw3.columns):
            AppNPS= (iw3.pivot_table(columns=['month_year'], 
                                    index=['AppType','fork'], 
                                    aggfunc=[np.mean,len], values='ratingvalue', margins=True)
                #.add_prefix('NPS_')
                #.reset_index()
            )
            AppNPSByType= (iw3.pivot_table(columns=['month_year'], 
                                    index=['AppType','UserType','fork'], 
                                    aggfunc=[np.mean,len], values='ratingvalue', margins=True)
                #.add_prefix('NPS_')
                #.reset_index()
                )
            AppNPS['mean'] = AppNPS['mean'].round(1)
            AppNPSByType['mean'] = AppNPSByType['mean'].round(1)

            AppNPS.rename(index=Builds, columns={'mean':'NPS','len':'Count'},inplace=True)
            AppNPSByType.rename(index=Builds, columns={'mean':'NPS','len':'Count'},inplace=True)

            AppNPS.to_csv(path+srq+app+'_Fork_Monthly'+d+'.csv')
            pd.DataFrame(data=['BY USER TYPE']).to_csv(app+'_Fork_Monthly'+d+'.csv',mode='a')
            AppNPSByType.to_csv(path+srq+app+'_Fork_Monthly'+d+'.csv',mode='a')
