import os
import csv
import math
import pandas as pd
import numpy as np
import datetime
import xlwt
import glob
import pyexcel as p

def transform_rating(rating):
    '''Input: Column of Data with NPS Field (on a scale of 1 to 5...)
    Output: Column with either Promoter/Detractor/Passive label, or the corresponding weights, based on datatype'''
    
    if rating == 5:
        return 'Promoter'
    elif rating <= 3:
        return 'Detractor'
    elif rating == 4:
        return 'Passive'
    else:
        return 'Invalid'
    

def NPS_rating(Rating_feel):
    if Rating_feel == 5.0:
        return 100
    elif Rating_feel== 4.0:
        return 0
    elif Rating_feel == 3.0:
        return -100
    elif Rating_feel == 2.0:
        return -100
    elif Rating_feel == 1.0:
        return -100
    else:
        return  0

def fork2(fork):
    "March Fork most Current"
    if fork <4:
        return 0
    else:
        return fork

def Sys_Vol_CatPercent(Percent):
    if Percent == 0:
        return '0-5%'
    elif Percent== 0.1:
        return '5-15%'
    elif Percent== 0.2:
        return '15-25%'
    elif Percent== 0.3:
        return '25-35%'
    elif Percent== 0.4:
        return '35-45%'
    elif Percent== 0.5:
        return '45-55%'
    elif Percent== 0.6:
        return '55-65%'
    elif Percent== 0.7:
        return '65-75%'
    elif Percent== 0.8:
        return '75-85%'
    elif Percent== 0.9:
        return '85-95%'
    elif Percent== 1.0:
        return '95-100%'


def create_NPSdataframe(data_file, build_file):
    '''Inputs: Raw Dataset and corresponding OS Builds
       Outputs: Cleaned Dataset'''
    iw3= pd.read_csv(data_file,index_col= 0)
    Build= pd.read_csv(build_file)

    iw3.columns = [x.partition('_')[2] for x in iw3.columns.to_list()]
    iw3 = iw3[list(iw3.columns[~iw3.columns.duplicated()])]

    iw3['minorbuild'] = iw3['OfficeBuild'].str.split('.', expand=True)[2]
    Build['OB_ThirdPart'] = Build['OfficeBuildPrefix'].str.split('.', expand=True)[2]
    Build =Build.set_index('OB_ThirdPart')['ForkBuild'].to_dict()
    iw3['fork'] = iw3['minorbuild'].astype(float).map(Build)
    iw3['fork'].fillna('<Other>',inplace=True)
    iw3['OfficeForkBuild'].fillna('Other',inplace=True)
    #Build =Build.set_index('OB_ThirdPart')['OfficeForkBuild'].to_dict()
    #iw3['Fork'] = iw3['minorbuild'].astype(float).map(Build)

        
    print('Mapped Builds to Forks')

    #iw3.rename(index=str,columns={"Source": "Product"},inplace=True)

    #iw3#.loc[iw['TenantState'] !='Unknown'] #This is a filter we sometimes use, but not currently
    # iw3=iw3.loc[iw3['App'] != 'Other']
    # iw3=iw3.loc[iw3['App'] != 'OneNote']
    # iw3=iw3[iw3['Product'] !='Desktop Visio']
    # iw3=iw3[iw3['Product'] !='Desktop OneNote']
    
    #iw3["SystemLocaleTag"] = iw3["SystemLocaleTag"].astype(str)
    #iw3["Language"] = iw3["SystemLocaleTag"].apply(lambda x: x.split("-")[0])

    iw3['fork2']=iw3['fork'].map(lambda x: fork2(x))

    iw3["LongDateTime"] = iw3["DateTime"].apply(lambda x: pd.to_datetime(x))
    #iw3= iw3.loc[iw3["LongDateTime"]>= datetime.strptime('2018-05-01' , '%Y-%m-%d').date()]

    iw3['Feed_year'] = pd.DatetimeIndex(iw3['LongDateTime']).year
    iw3['Feed_month'] = pd.DatetimeIndex(iw3['LongDateTime']).month
    iw3['month_year'] = iw3.LongDateTime.dt.to_period('M')
    iw3['week_year'] = iw3.LongDateTime.dt.to_period('W-SAT')
    iw3['Feed_week'] = pd.DatetimeIndex(iw3['LongDateTime']).week
    print('Analysis dates: {0} to {1}'.format(iw3["LongDateTime"].min().strftime('%Y-%m-%d'), iw3["LongDateTime"].max().strftime('%Y-%m-%d')))

    iw3['NPS'] = iw3['Rating'].apply(transform_rating)
    iw3['ratingvalue'] = iw3['Rating'].apply(NPS_rating)

    # iw3.loc[iw3['SurveyRatingQuestion'] == 'How likely are you to recommend Office 365 to a friend or colleague?', 'SurveyType'] = 'Suite'
    # iw3.loc[iw3['SurveyRatingQuestion'] == 'How likely are you to recommend our application to a friend or colleague?', 'SurveyType'] = 'App'    
    iw3.loc[iw3['SurveyName'].str.startswith('Suite'), 'SurveyType'] = 'Suite'
    iw3.loc[iw3['SurveyName'].str.startswith('App'), 'SurveyType'] = 'App'
    iw3['AppType'] = iw3['SurveyType'] + iw3['App']
    
    #Redsigning Discretized Variables:
    iw3.loc[iw3['ScreenDpi']<120,'ScreenDPI'] = 'Less than 120'
    iw3.loc[iw3['ScreenDpi']>=120,'ScreenDPI'] = '120 or Greater'
    iw3['RamGB'] = iw3['RamMB']/1024.0
    iw3['ProcSpeedGHz'] = iw3['ProcSpeedMHz']//1000
    
    count_list = [#'TotalUsers'#,'ProPlusEnabledUsers'
    ]
    
    iw3.loc[iw3['SysVolSizeMB']>=900000,'SysVolSizeGB'] = 'A Tier (~1TB or More)'
    iw3.loc[iw3['SysVolSizeMB'].between(450000,900000,inclusive=False),'SysVolSizeGB'] = 'B Tier (~0.5TB to ~1TB)'
    iw3.loc[iw3['SysVolSizeMB'].between(250000,450000,inclusive=True),'SysVolSizeGB'] = 'C Tier (~256 to ~500GB)'
    iw3.loc[iw3['SysVolSizeMB'].between(125000,250000,inclusive=False),'SysVolSizeGB'] = 'D Tier (~128 to ~256GB)'
    iw3.loc[iw3['SysVolSizeMB'].between(64000,128000,inclusive=True),'SysVolSizeGB'] = 'E Tier (~64 to ~128GB)'
    iw3.loc[iw3['SysVolSizeMB']<64000,'SysVolSizeGB'] = 'F Tier (Less than 64GB)'

    iw3.loc[iw3['SysVolFreeSpaceMB']>500000,'SysVolFreeSpaceGB'] = 'Most Free Space (More Than 0.5TB)'
    iw3.loc[iw3['SysVolFreeSpaceMB'].between(100000,500000,inclusive=True),'SysVolFreeSpaceGB'] = 'Moderate Free Space (100GB to 500GB)'
    iw3.loc[iw3['SysVolFreeSpaceMB']<100000,'SysVolFreeSpaceGB'] = 'Least Free Space (Less Than 100GB)'
    
    iw3['SysVolUsagePercent'] = 1-iw3['SysVolFreeSpaceMB']/iw3['SysVolSizeMB']
    iw3['SysVolUsagePercent'] = iw3['SysVolUsagePercent'].apply(lambda x: round(x,1))
    iw3['SysVolUsagePercent'] = iw3['SysVolUsagePercent'].apply(Sys_Vol_CatPercent)

    for i, item in enumerate(count_list):
        var_name = item+'_Ord'
        iw3.loc[iw3[item]==1,var_name] = '1'
        iw3.loc[iw3[item].between(2,9),var_name] = 'Small Group (2-9)'
        iw3.loc[iw3[item].between(10,149),var_name] = 'Medium Group (10-149)'
        iw3.loc[iw3[item]>=150,var_name] = 'Large Group (150+)'
    
    iw3['Resolution']=iw3['HorizontalResolution'].astype(str)+'x'+iw3['VerticalResolution'].astype(str)
    iw3[iw3.columns[list(iw3.dtypes=='category')]] = iw3[iw3.columns[list(iw3.dtypes=='category')]].astype(str)
    
    return iw3

# df_S = create_NPSdataframe('EndUserSuite(April_May).csv','../Downloads/Build_to_fork_May2019.csv')
# df_A = create_NPSdataframe('EndUserApp(April_May).csv','../Downloads/Build_to_fork_May2019.csv')
# df = pd.concat([df_S.loc[:,df_S.columns==df_A.columns],df_A.loc[:,df_S.columns==df_A.columns]])
d = '('+str(datetime.datetime.today().month)+'-'+str(datetime.datetime.today().day)+'-'+str(datetime.datetime.today().year)+')'
# print(d)


def make_pivot_table_new(df, interval,previous, current, value, dim=None):
    if dim == None: #AllUp Table, with No Splits by Levels of another Variable
        table = df.pivot_table(columns=[interval], aggfunc=[np.mean, len], values=value).dropna()
        S_previous = sum(table[('len',previous)]) 
        S_current = sum(table[('len',current)])
        table[('Sample',previous)] = table[('len',previous)]/S_previous
        table[('Sample',current)] = table[('len',current)]/S_current
        table[('Impact','SampleImpact')] = table[('mean',previous)]*(table[('Sample',current)]- table[('Sample',previous)])
        table[('Impact','TrueNPSImpact')] = table[('Sample',current)]*(table[('mean',current)]- table[('mean', previous)])
        table[('Impact','TotalImpact')] = table[('Impact','SampleImpact')] + table[('Impact','TrueNPSImpact')]
        table[('Test', 'MoE')] = 2*100*((abs(table[('mean',previous)]/100)*(1-abs(table[('mean',previous)]/100)))/(table[('len',previous)]-1))**0.5
        table[('Test', 'StatSig')] = abs((table[('Impact','TrueNPSImpact')])) > table[('Test','MoE')]
        #table = table[table.len.min(axis=1)>30]
        
    if dim != None: #Split by levels of another variable
        table = df.pivot_table(columns=[interval], index=[dim], aggfunc=[np.mean, len], values=value).dropna()
        S_previous = sum(table[('len',previous)]) 
        S_current = sum(table[('len',current)])
        table[('Sample',previous)] = table[('len',previous)]/S_previous
        table[('Sample',current)] = table[('len',current)]/S_current
        table[('Impact','SampleImpact')] = table[('mean',previous)]*(table[('Sample',current)]- table[('Sample',previous)])
        table[('Impact','TrueNPSImpact')] = table[('Sample',current)]*(table[('mean',current)]- table[('mean', previous)])
        table[('Impact','TotalImpact')] = table[('Impact','SampleImpact')] + table[('Impact','TrueNPSImpact')]
        table[('Test', 'MoE')] = 2*100*((abs(table[('mean',previous)]/100)*(1-abs(table[('mean',previous)]/100)))/(table[('len',previous)]-1))**0.5
        table[('Test', 'StatSig')] = abs((table[('Impact','TrueNPSImpact')])) > table[('Test','MoE')]
        table = table[table.len.min(axis=1)>30]
    return table.round(4)

# def make_pivot_table_new(df, interval,previous, current, value, dim=None):
#     if dim == None: #AllUp Table, with No Splits by Levels of another Variable
#         table = df.pivot_table(columns=[interval], aggfunc=[np.mean, len], values=value).dropna()
#         S_previousAll = sum(table[('len',previous)]) 
#         S_currentAll = sum(table[('len',current)])
#         table[('Sample',previous)] = table[('len',previous)]/S_previousAll
#         table[('Sample',current)] = table[('len',current)]/S_currentAll
#         sample_impact1 = table[('mean',previous)]*(table[('Sample',current)]- table[('Sample',previous)])
#         sample_impact2 = table[('mean',current)]*(table[('Sample',current)]- table[('Sample',previous)])
#         nps_impact1 = table[('Sample',current)]*(table[('mean',current)]- table[('mean', previous)])
#         nps_impact2 = table[('Sample',previous)]*(table[('mean',current)]- table[('mean', previous)])
#         table[('Impact','SampleImpact')] = (sample_impact1+sample_impact2)/2
#         table[('Impact','TrueNPSImpact')] = (nps_impact1+nps_impact2)/2
#         table[('Impact','TotalImpact')] = table[('Impact','SampleImpact')] + table[('Impact','TrueNPSImpact')]
#         table[('Test', 'MoE')] = 2*100*((abs(table[('mean',previous)]/100)*(1-abs(table[('mean',previous)]/100)))/(table[('len',previous)]-1))**0.5
#         if table[('Test', 'MoE')][0] ==0:
#             table[('Test', 'MoE')] = 2*100*((abs(table[('mean',current)]/100)*(1-abs(table[('mean',current)]/100)))/(table[('len',previous)]-1))**0.5

#         table[('Test', 'StatSig')] = abs((table[('Impact','TrueNPSImpact')])) > table[('Test','MoE')]
#         table = table[table.len.min(axis=1)>30]
#         table.rename({'mean':'NPS','len':'Count'},axis='columns',inplace=True)
        
#     if dim != None: #Split by levels of another variable
#         table = df.pivot_table(columns=[interval], index=[dim], aggfunc=[np.mean, len], values=value).dropna()
#         S_previous = sum(table[('len',previous)]) 
#         S_current = sum(table[('len',current)])
#         table[('Sample',previous)] = table[('len',previous)]/S_previous
#         table[('Sample',current)] = table[('len',current)]/S_current
#         sample_impact1 = table[('mean',previous)]*(table[('Sample',current)]- table[('Sample',previous)])
#         sample_impact2 = table[('mean',current)]*(table[('Sample',current)]- table[('Sample',previous)])
#         nps_impact1 = table[('Sample',current)]*(table[('mean',current)]- table[('mean', previous)])
#         nps_impact2 = table[('Sample',previous)]*(table[('mean',current)]- table[('mean', previous)])
#         table[('Impact','SampleImpact')] = (sample_impact1+sample_impact2)/2
#         table[('Impact','TrueNPSImpact')] = (nps_impact1+nps_impact2)/2
#         table[('Impact','TotalImpact')] = table[('Impact','SampleImpact')] + table[('Impact','TrueNPSImpact')]
#         table[('Test', 'MoE')] = 2*100*((abs(table[('mean',previous)]/100)*(1-abs(table[('mean',previous)]/100)))/(table[('len',previous)]-1))**0.5
#         print(table[('Test','MoE')])
#         for x in table[('Test', 'MoE')].values:
#             #print(x)
#             if x == 0:
#                 x = 2*100*((abs(table[('mean',current)]/100)*(1-abs(table[('mean',current)]/100)))/(table[('len',previous)]-1))**0.5
#         table[('Test', 'StatSig')] = abs((table[('Impact','TrueNPSImpact')])) > table[('Test','MoE')]
#         table = table[table.len.min(axis=1)>30]
#     table.rename({'mean':'NPS','len':'Count'},axis='columns',inplace=True)
#     return table.round(4)



def custom_impact_pivot_allup(df, time_interval,previous,current):
    '''Inputs: Data, suite and app, and consecutive months - in 'YYY-MM' format- 
    and signals is a [LIST] of variables for which we want to examine impact/contribution
    
    DO NOT RUN if data does not have Suite Survey, you will get a (len,prev) error 
    This is because you will be working with a dataset that has filtered out all non-suite
    i.e. an empty dataset'''

    if 'Custom_Period' in df.columns:
        df.drop('Custom_Period',axis=1,inplace=True)
    
    if time_interval.endswith('_year'):
        previous_periods = [pd.Period(x) for x in previous]
        current_periods = [pd.Period(y) for y in current]
    else:
        previous_periods = previous.copy()
        current_periods = current.copy()
    
    previous = ''.join(str(previous))
    current = ''.join(str(current))
    
    stamp = time_interval.split('_')[0].upper()
    #stamp = stamp[0]
    #print(stamp)
    print(previous_periods,current_periods)
    
    
    common_list = ["OfficeForkBuild","Feedback_Month","A13Region","OfficeUiLanguage","OsBuild", "OfficeBuild",#"Language",
                   "DeviceManufacturer","OfficeArchitectureText","RamGB","ProcSpeedGHz", 'SysVolUsagePercent',
                   "SysVolSizeGB","SysVolFreeSpaceGB","Resolution","ScreenDPI","InstallType"]

    #commerce_list = ["Tenant_Tenure_UHICommercial","Usage_Segment_UHICommercial","HasEducation_UHICommercial",
    #                "FinalSegment_UHICommercial","FinalSubsegment_UHICommercial","SeatSizeBuckets_UHICommercial"]
    #also includes TenantState

    # consumer_list = ["Consumer_TenureBucket", "Consumer_LicenseType","Consumer_LicenseType",
    #                 "EP_SubscriptionStatus_UHIConsumer","UsageSegment_v2_UHIConsumer",
    #                  "UserStage_UHIConsumer","UserStageName_UHIConsumer","EP_SKUCategory_UHIConsumer"]
    commerce_list = []

    for com in list(df.columns[df.columns.str.startswith('Commercial')]):
        if df[com].nunique() >=2 and df[com].nunique()<10:
            commerce_list.append(com)

    consumer_list = []

    for con in list(df.columns[df.columns.str.startswith('Consumer')]):
        if df[con].nunique() >=2 and df[con].nunique()<10:
            consumer_list.append(con)


    if time_interval.startswith('fork'):
        #common_list.remove('fork')
        common_list.remove('OfficeForkBuild')

    if time_interval.startswith('month'):
        common_list.remove('Feedback_Month')

    make_missing = ['ProcessorCount'#, 'ProductName_OdinConsumer','LicenseType','TenantState'
    ]

    print(df[time_interval].max())
    for p in previous_periods:
        print(p)
        df.loc[df[time_interval]==p,'Custom_Period'] = previous
        print(df['Custom_Period'].unique())
    #df.loc[df[time_interval].isin(previous_periods),'Custom_Period'] = previous
    
    for c in current_periods:
        df.loc[df[time_interval]==c,'Custom_Period'] = current
    #df.loc[df[time_interval].isin(current_periods),'Custom_Period'] = current
    
    print(df.shape)
    # df = df[#(df.A13Region.notnull())# & (df.OsBuild.isin(builds)) \
    #          & (df.OsName == 'Windows 10')\
    #              & (df.SurveyType=='Suite')
    #          & (df.OlsOfferType == 'Subscription')\
    #          & (df.OsName == 'Windows10')\
    #          & (df.Platform == 'Windows Desktop')\
    #          & (df.Channel == 'CC') \
    #          & (df.AudienceGroup == 'Production')
    #        ] 
    print(df.shape)
    df.loc[:,make_missing] = df.loc[:,make_missing].fillna('Missing')
    common_list.extend(make_missing)
    common_list.insert(0,'App')
    
    print(df.Custom_Period.unique(), previous, current)
    
    table_0 = make_pivot_table_new(df,'Custom_Period',previous,current,value='ratingvalue')
    table_0.rename({'ratingvalue': 'AllUp'}, axis='index',inplace=True)
    table_1 = make_pivot_table_new(df,'Custom_Period',previous,current,value='ratingvalue',dim='UserType')
    
    editions = df.UserType.dropna().unique()
    for edition in editions:
        slices = common_list.copy()
        #print(signals)
        if edition =="Commercial":
            slices.extend(commerce_list)
            #slices.remove('LicenseType')
        elif edition =="Consumer":
            slices.extend(consumer_list)
            #slices.remove('TenantState')
        data = df[(df.UserType==edition) & (df.SurveyType=='Suite')]
        #print(data.pivot_table(columns=['Custom_Period'],values='ratingvalue'))
        table_2 = make_pivot_table_new(data,'Custom_Period',previous,current,value='ratingvalue')
        table_2.rename({'ratingvalue': 'AllUp'+edition}, axis='index',inplace=True)
        df_empty = pd.DataFrame({' ' : []})
        table_0.to_csv('SuiteAllUp_Impact'+edition[:4]+stamp+'.csv',)
        table_1[table_1.index==edition].to_csv('SuiteAllUp_Impact'+edition[:4]+stamp+'.csv',header=False,mode='a')
        table_2.to_csv('SuiteAllUp_Impact'+edition[:4]+stamp+'.csv',header=False,mode='a')
        df_empty.to_csv('SuiteAllUp_Impact'+edition[:4]+stamp+'.csv',mode='a')
        #print(edition, slices)
        for var in slices:
            table = make_pivot_table_new(data,'Custom_Period',previous,current,dim=var,value='ratingvalue')
            if var == 'DeviceManufacturer':
                table = table.query('DeviceManufacturer != ["To Be Filled By O.E.M."]')
            elif var == 'Resolution':
                table = table.query('Resolution != ["nanxnan"]')
            if var in ['A13Region', 'DeviceManufacturer']:
                table = table.sort_values(by=[('Impact','TotalImpact')])
            elif var in ["Tenant_Tenure_UHICommercial","SeatSizeBuckets_UHICommercial"]:
                table = table.sort_index()
            pd.DataFrame(data=[var]).to_csv('SuiteAllUp_Impact'+edition[:4]+stamp+'.csv',mode='a',header=False,index=False)
            table.to_csv('SuiteAllUp_Impact'+edition[:4]+stamp+'.csv',header=False,mode='a')
            df_empty.to_csv('SuiteAllUp_Impact'+edition[:4]+stamp+'.csv',mode='a')

def custom_impact_pivot_levels(df,time_interval, s_o_a,
                         previous,current):
    '''Inputs: Data, suite and app, and consecutive months in 'YYYY-MM' format,
    and signals is a [LIST] of variables for which we want to examine impact/contribution'''
    #periods = [pd.Period(previous),pd.Period(current)] #Specifying the months for comparison
    
    if 'Custom_Period' in df.columns:
        df.drop('Custom_Period',axis=1,inplace=True)
    
    if time_interval.endswith('_year'):
        previous_periods = [pd.Period(x) for x in previous]
        current_periods = [pd.Period(y) for y in current]
    else:
        previous_periods = previous.copy()
        current_periods = current.copy()
    
    previous = ''.join(str(previous))
    current = ''.join(str(current))
    
    stamp = time_interval.split('_')[0].upper()
    #stamp = stamp[0]
    print(stamp)
    common_list = ["OfficeForkBuild","Feedback_Month","A13Region","OfficeUiLanguage","OsBuild", "OfficeBuild",#"Language",
                   "DeviceManufacturer","OfficeArchitectureText","RamGB","ProcSpeedGHz", 'SysVolUsagePercent',
                   "SysVolSizeGB","SysVolFreeSpaceGB","Resolution","ScreenDPI","InstallType"]

    #commerce_list = ["Tenant_Tenure_UHICommercial","Usage_Segment_UHICommercial","HasEducation_UHICommercial",
    #                "FinalSegment_UHICommercial","FinalSubsegment_UHICommercial","SeatSizeBuckets_UHICommercial"]
    #also includes TenantState

    # consumer_list = ["Consumer_TenureBucket", "Consumer_LicenseType","Consumer_LicenseType",
    #                 "EP_SubscriptionStatus_UHIConsumer","UsageSegment_v2_UHIConsumer",
    #                  "UserStage_UHIConsumer","UserStageName_UHIConsumer","EP_SKUCategory_UHIConsumer"]
    commerce_list = []

    for com in list(df.columns[df.columns.str.startswith('Commercial')]):
        if df[com].nunique() >=2 and df[com].nunique()<10:
            commerce_list.append(com)

    consumer_list = []

    for con in list(df.columns[df.columns.str.startswith('Consumer')]):
        if df[con].nunique() >=2 and df[con].nunique()<10:
            #print(con,'IS ADDED') Consumer fields are all null as of October 23
            consumer_list.append(con)
    print('Consumer List is changing:', consumer_list)
    
    print(previous)

    #also includes LicenseType
    
    if time_interval.startswith('fork'):
        common_list.remove('OfficeForkBuild')

    if time_interval.startswith('month'):
        common_list.remove('Feedback_Month')        
    
    make_missing = ['ProcessorCount',#'ProductName_OdinConsumer'#,'LicenseType','TenantState'
    ]

    print(df[time_interval].max())
    for p in previous_periods:
        print('Previous Period ',p)
        df.loc[df[time_interval]==p,'Custom_Period'] = previous
    #df.loc[df[time_interval].isin(previous_periods),'Custom_Period'] = previous
    
    for c in current_periods:
        print('Current Period ',c)
        df.loc[df[time_interval]==c,'Custom_Period'] = current
    #df.loc[df[time_interval].isin(current_periods),'Custom_Period'] = current
    
    print(df.shape)
    # df = df[(df.OsName == 'Windows 10')# & (df.OsBuild.isin(builds)) \
    #          & (df.A13Region.notnull())\
    #              & (df.SurveyType=='Suite')
    #          & (df.OlsOfferType == 'Subscription')\
    #          & (df.OsName == 'Windows10')\
    #          & (df.Platform == 'Windows Desktop')\
    #          & (df.Channel == 'CC') \
    #          & (df.AudienceGroup == 'Production')
    #        ] 
    # print(df.shape)
    common_list.extend(make_missing)
    
    print(df.dropna().Custom_Period.unique(), previous, current)
    #df = df[~((df.Custom_Period.isin(['2019-01_2019-02'])) & (df.OsBuild.isin(['10.0.16299'])))]
    #df = df[~((df.Custom_Period.isin([current])) & (df.OsBuild.isin([builds[lag-1]])))] 
    
    #     all_up_table_0 = make_pivot_table_new(df,'Custom_Period',previous,current,'ratingvalue')
    #     all_up_table_0.index = ['AllUp']
    #     all_up_table_0.to_csv('Outlook_ImpactCustom.csv')
    
    df_empty = pd.DataFrame({' ' : []})
    #     df_empty.to_csv('Outlook_ImpactCustom.csv',mode='a')
        
    #     all_up_table_ed = make_pivot_table_new(df,'Custom_Period',previous,current,
    #                                           'ratingvalue',dim='Office365Category_ACID')
    #     pd.DataFrame(data=['Office365Category_ACID']).to_csv('Outlook_ImpactCustom.csv',mode='a',header=False,index=False)
    #     all_up_table_ed.to_csv('Outlook_ImpactCustom.csv', mode='a',header=False)
    #     df_empty.to_csv('Outlook_ImpactCustom.csv',mode='a')
    
    #Overall App Logic
    Apps = df.App.dropna().unique()
    for app in Apps:
        dta = df[(df.App==app) & (df.SurveyType==s_o_a)]
        print(dta.shape, s_o_a, df.SurveyType.unique())
        table_0 = make_pivot_table_new(dta,'Custom_Period', previous,current, value='ratingvalue')
        table_0.index = ['Overall '+app]
        table_1 = make_pivot_table_new(dta,'Custom_Period', previous,current, value='ratingvalue', dim = 'UserType')
        table_0.to_csv(s_o_a+app+'All_Impact'+stamp+'.csv')
        table_1.to_csv(s_o_a+app+'All_Impact'+stamp+'.csv',header=False,mode='a')
        df_empty.to_csv(s_o_a+app+'All_Impact'+stamp+'.csv',mode='a')
        for var in common_list:
                print(var)
                table = make_pivot_table_new(dta,'Custom_Period',previous,current,'ratingvalue',dim=var)
                if var == 'DeviceManufacturer':
                    table = table.query('DeviceManufacturer != ["To Be Filled By O.E.M."]')
                elif var == 'Resolution':
                    table = table.query('Resolution != ["nanxnan"]')
                if var in ['A13Region', 'DeviceManufacturer']:
                    table.reindex(table['Impact'].sort_values(by='TotalImpact', ascending=True).index)
                elif var in []:
                    table = table.sort_index()
                pd.DataFrame(data=[var]).to_csv(s_o_a+app+'All_Impact'+stamp+'.csv',mode='a',header=False,index=False)
                table.to_csv(s_o_a+app+'All_Impact'+stamp+'.csv',header=False,mode='a')
                df_empty.to_csv(s_o_a+app+'All_Impact'+stamp+'.csv',mode='a')
    
    editions = df.UserType.dropna().unique()
    Survey_Apps = df.AppType.dropna().unique()
    for edition in editions:
        slices = common_list.copy()
        #print(signals)
        if edition =="Commercial":
            slices.extend(commerce_list)
            #slices.remove('LicenseType')
        #elif edition =="Consumer":
            #slices.extend(consumer_list)
            #slices.remove('TenantState')
        for level in Survey_Apps:
            data = df[(df.UserType==edition)&(df.AppType==level)]
            #print(level,edition)
            #print(data.shape)
            all_up_table_0 = make_pivot_table_new(data,'Custom_Period',previous,current,'ratingvalue',dim='AppType')
            all_up_table_0.index = [level+edition]
            all_up_table_0.to_csv(''+(level.replace('_','')+'_Impact_'+edition[:4]+stamp)[:32]+'.csv')
            df_empty.to_csv(''+(level.replace('_','')+'_Impact_'+edition[:4]+stamp)[:32]+'.csv', mode='a')
            for var in slices:
                print(var)
                table = make_pivot_table_new(data,'Custom_Period',previous,current,'ratingvalue',dim=var)
                if var == 'DeviceManufacturer':
                    table = table.query('DeviceManufacturer != ["To Be Filled By O.E.M."]')
                elif var == 'Resolution':
                    table = table.query('Resolution != ["nanxnan"]')
                if var in ['A13Region', 'DeviceManufacturer']: #i.e. if we dont care about alphabetical order index
                    table.reindex(table['Impact'].sort_values(by='TotalImpact', ascending=True).index)
                elif var in []:
                    table = table.sort_index()
                pd.DataFrame(data=[var]).to_csv(''+(level.replace('_','')+'_Impact_'+edition[:4]+stamp)[:32]+'.csv'
                                                ,mode='a',header=False,index=False)
                table.to_csv(''+(level.replace('_','')+'_Impact_'+edition[:4]+stamp)[:32]+'.csv',
                             header=False,mode='a')
                df_empty.to_csv(''+(level.replace('_','')+'_Impact_'+edition[:4]+stamp)[:32]+'.csv',mode='a')
        #print(table, data.shape, df.shape)

def verbatim_topic_counts(df, interval,previous,current,dictionary):
    '''Apply Classification Tags before running this function on the dataset.
    Specify interval of interest (Month,Fork,Week): for fork example: [10,11,12]'''
    #verbatim_classification_from_dict(dictionary, df)
    final_df = pd.DataFrame()
    
    if interval.endswith('_year'):
        previous_periods = [pd.Period(x) for x in previous]
        current_periods = [pd.Period(y) for y in current]
        previous = '_'.join(str(previous))
        current = '_'.join(str(current))
    else:
        previous_periods = previous.copy()
        current_periods = current.copy()
        previous = ''.join(str(previous))
        current = ''.join(str(current))
    
    df.loc[df[interval].isin(previous_periods),'Custom_Period'] = previous
    df.loc[df[interval].isin(current_periods),'Custom_Period'] = current
        

    for i in [previous,current]:
        #calculate total occurences for each topic, then append the total number of verbatims at the bottom
        df_i = df[(df['Custom_Period'] == i)]
        PromotorVerbatims = df_i[df_i.Rating==5].iloc[:,-len(dictionary)-1:-1]
        DetractorVerbatims = df_i[df_i.Rating.isin([1,2,3])].iloc[:,-len(dictionary)-1:-1]
        PromotorSums = PromotorVerbatims.sum()
        DetractorSums = DetractorVerbatims.sum()
        VCounts = pd.concat([PromotorSums,DetractorSums],axis=1)
        VCounts.loc['TotalVerbatims'] = [len(PromotorVerbatims),len(DetractorVerbatims)]
        #VCounts.loc['TotalVerbatims'] = [len(PromotorVerbatims),len(DetractorVerbatims)]
        VCounts.columns = ['Promotors'+i,'Detractors'+i]
        VCounts.index = VCounts.index.map(lambda x: x.lstrip('verbatim_'))
        final_df = pd.concat([final_df,VCounts],axis=1)
    percent_df = final_df.divide(final_df.iloc[-1]).round(4)* 100
    percent_df.iloc[-1] = final_df.iloc[-1]
    return final_df , percent_df


# time_interval = 'month_year' # or fork or week_year...
# stamp = time_interval.split('_')[0].upper()

# num_check = set('0123456789.')
# letter_check = set('qwertyuiopasdfghjklzxcvbnm-+()[]{}')

# wb = xlwt.Workbook()
# for filename in glob.glob("Suite*"+stamp+".csv"):
#     (f_path, f_name) = os.path.split(filename)
#     (f_short_name, f_extension) = os.path.splitext(f_name)
#     print(f_short_name)
#     ws = wb.add_sheet(f_short_name.partition(stamp)[0])
#     fileReader = csv.reader(open(filename, 'rt'))
#     for rowx, row in enumerate(fileReader):
#         for colx, value in enumerate(row):
#             if (any((n in value) for n in num_check)) & ('.0.' not in value) & (any((c in value.lower()) for c in letter_check) is False):
#                 num_value = float(value)
#                 ws.write(rowx, colx, num_value)
#             elif value.strip().startswith('-'):
#                 num_value = float(value)
#                 ws.write(rowx, colx, num_value)
#             else:
#                 ws.write(rowx, colx, value)
                
                    
# wb.save("SuiteImpactReports"+stamp+d+".xls")
# p.save_book_as(file_name='SuiteImpactReports'+stamp+d+'.xls',
#                dest_file_name='SuiteImpactReports'+stamp+d+'.xlsx')

# wb = xlwt.Workbook()
# for filename in glob.glob("App*"+stamp+".csv"):
#     (f_path, f_name) = os.path.split(filename)
#     (f_short_name, f_extension) = os.path.splitext(f_name)
#     ws = wb.add_sheet(f_short_name.partition(stamp)[0])
#     fileReader = csv.reader(open(filename, 'rt'))
#     for rowx, row in enumerate(fileReader):
#         for colx, value in enumerate(row):
#             if (any((n in value) for n in num_check)) & ('.0.' not in value) & (any((c in value.lower()) for c in letter_check) is False):
#                 #print(value)
#                 num_value = float(value)
#                 ws.write(rowx, colx, num_value)
#             elif value.strip().startswith('-'):
#                 num_value = float(value)
#                 ws.write(rowx, colx, num_value)
#             else:
#                 ws.write(rowx, colx, value)
                
                    
# wb.save("AppImpactReports"+stamp+d+".xls")
# p.save_book_as(file_name='AppImpactReports'+stamp+d+'.xls',
#                dest_file_name='AppImpactReports'+stamp+d+'.xlsx')