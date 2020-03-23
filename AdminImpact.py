import os
import csv
import math
import pandas as pd
import numpy as np
import datetime

#Script by Jared Anderson

def NPS_rating(rating):
    '''Input: Column of Data with NPS Field (on a scale of 1 to 5...)
    Output: Column with either Promoter/Detractor/Passive label, or the corresponding weights, based on datatype'''
    if rating == 5:
        return 100
    elif rating <= 3:
        return -100
    elif rating == 4:
        return 0
    else:
        return np.nan
    
d = '('+str(datetime.datetime.today().month)+'-'+str(datetime.datetime.today().day)+'-'+str(datetime.datetime.today().year)+')'
def create_NPSdataframe(data_file):
    df = pd.read_csv(data_file, low_memory=True)
    df['NPS'] = df['Feedback_Rating'].apply(NPS_rating)
    ######################################################################
    df.Tenant_MSSalesAreaName = df.Tenant_MSSalesAreaName.fillna('Other')#
    ######################################################################
    ###################################################################
    make_missing = ['Tenant_TenantCategory','Tenant_CustomerSegmentGroup',
                    'Tenant_MSSalesAreaName','LastTicket_ProgramName','Admin_UsageSegment',
                'LastTicket_Source','LastTicket_PremierOffering']

    df.loc[:,make_missing] = df.loc[:,make_missing].fillna('Missing')#
    ####################################################################
    df["LongDateTime"] = df["Feedback_DateTime"].apply(lambda x: pd.to_datetime(x))
    df['Feed_year'] = pd.DatetimeIndex(df['LongDateTime']).year
    df['Feed_month'] = pd.DatetimeIndex(df['LongDateTime']).month
    df['month_year'] = df.LongDateTime.dt.to_period('M')
    df['week_year'] = df.LongDateTime.dt.to_period('W-SAT')
    df['Feed_week'] = pd.DatetimeIndex(df['LongDateTime']).week

    ##################################################################
    visit_list = ['Admin_PortalVisits_15Days','Admin_PortalVisits_30Days']

    for i, item in enumerate(visit_list):
        var_name = item+'_Ord'
        df.loc[df[item].between(0,1),var_name] = '1 Visit At Most'
        df.loc[df[item].between(2,4),var_name] = 'Least Visits (2-4)'
        df.loc[df[item].between(5,9),var_name] = 'Moderate Visits (5-9)'
        df.loc[df[item]>=10,var_name] = 'Most Visits (10+)'
    ##################################################################

    count_list = [#'Tenant_AssignedPlanCount','Tenant_TotalGroupCount',
                #'Tenant_Odin_LicensedUsers','Tenant_Odin_TotalUsers',
                'Tenant_TotalSubscriptionsCount','Tenant_TotalUsers']

    for i, item in enumerate(count_list):
        var_name = item+'_Ord'
        df.loc[df[item]==1,var_name] = 'Single User'
        df.loc[df[item].between(2,9),var_name] = 'Group - Small (2-9)'
        df.loc[df[item].between(10,149),var_name] = 'Group - Medium (10-149)'
        df.loc[df[item]>=150,var_name] = 'Group - Large (150+)'
    # ##################################################################
    # def time_split(df,col,quantiles=True,day_groups=25,hour_groups=10):
    #     day_name = col+'Days'
    #     hour_name = col+'Hours'
    #     min_name = col+'Mins'
    #     day_else = df[col].str.rpartition('.',expand=True).iloc[:,[0,2]]
    #     hour_else = day_else[2].str.partition(':')
    #     minute_else = hour_else[2].str.partition(':')
    #     df[day_name] = pd.to_numeric(day_else[0]) 
    #     df[hour_name] = pd.to_numeric(hour_else[0]) #if Days is 0
    #     df[min_name] = pd.to_numeric(minute_else[0]) # if Days & Hours are 0
    #     m = df[day_name].notnull()
    #     n = df[hour_name]>=1
    #     df.loc[:,[hour_name,min_name]] = df.loc[:,[hour_name,min_name]].mask(m, np.nan)
    #     df.loc[:,min_name] = df.loc[:,min_name].mask(n, np.nan)
    #     if quantiles==True:
    #         df[day_name+'Ord'] = pd.qcut(df[day_name],day_groups,duplicates='drop')
    #         df[hour_name+'Ord'] = pd.qcut(df[hour_name],hour_groups,duplicates='drop')
    #         df[min_name+'Ord'] = df[min_name].round(-1)
    #     #day_list.append(day_name), hour_list.append(hour_name), minute_list.append(min_name)
    # ##################################################################
    # time_split(df,'LastTicket_TimeToResolve')
    # time_split(df,'LastTicket_TimeToLastTicket',day_groups=9,hour_groups=2)

    ##################################################################
    verbatim_list = list(df.columns[df.columns.str.startswith('Verbatim_Sum_')])
    ##################################################################
    ##################################################################
    for v in verbatim_list:
        df[v][df.loc[:,v]>=1]=1
    ##################################################################
    ##################################################################
    subscription_list = list(df.columns[df.columns.str.startswith('Subscription_Sum_')])
    ##################################################################
    ##################################################################
    for v in subscription_list:
        df[v][df.loc[:,v]>=1]=1
    ##################################################################
    df[df.columns[list(df.dtypes=='category')]] = df[df.columns[list(df.dtypes=='category')]].astype(str)
    return df

def make_admin_pivot_new(df, interval,previous, current,dim=None):
    '''For calling WITHIN the Impact Report function, so that df is pre-filtered by the periods of interest'''
    if dim == None:
        
        all_up = df.pivot_table(columns=[interval], aggfunc=[np.mean, len], values='NPS').dropna()
        S_previous = sum(all_up[('len',previous)]) 
        S_current = sum(all_up[('len',current)])
        all_up[('Sample',previous)] = all_up[('len',previous)]/S_previous
        all_up[('Sample',current)] = all_up[('len',current)]/S_current
        all_up[('Impact','SampleImpact')] = all_up[('mean',previous)]*(all_up[('Sample',current)]- all_up[('Sample',previous)])
        all_up[('Impact','TrueNPSImpact')] = all_up[('Sample',current)]*(all_up[('mean',current)]- all_up[('mean', previous)])
        all_up[('Impact','TotalImpact')] = all_up[('Impact','SampleImpact')] + all_up[('Impact','TrueNPSImpact')]
        all_up[('Test', 'MoE')] = 2*100*((abs(all_up[('mean',previous)]/100)*(1-abs(all_up[('mean',previous)]/100)))/(all_up[('len',previous)]-1))**0.5
        all_up[('Test', 'StatSig')] = abs((all_up[('Impact','TrueNPSImpact')])) > all_up[('Test','MoE')]
        all_up = all_up[all_up.len.min(axis=1)>30]
        
    else:
        all_up = df.pivot_table(columns=[interval], index=[dim], aggfunc=[np.mean, len], values='NPS').dropna()
        S_previous = sum(all_up[('len',previous)]) 
        S_current = sum(all_up[('len',current)])
        all_up[('Sample',previous)] = all_up[('len',previous)]/S_previous
        all_up[('Sample',current)] = all_up[('len',current)]/S_current
        all_up[('Impact','SampleImpact')] = all_up[('mean',previous)]*(all_up[('Sample',current)]- all_up[('Sample',previous)])
        all_up[('Impact','TrueNPSImpact')] = all_up[('Sample',current)]*(all_up[('mean',current)]- all_up[('mean', previous)])
        all_up[('Impact','TotalImpact')] = all_up[('Impact','SampleImpact')] + all_up[('Impact','TrueNPSImpact')]
        all_up[('Test', 'MoE')] = 2*100*((abs(all_up[('mean',previous)]/100)*(1-abs(all_up[('mean',previous)]/100)))/(all_up[('len',previous)]-1))**0.5
        all_up[('Test', 'StatSig')] = abs((all_up[('Impact','TrueNPSImpact')])) > all_up[('Test','MoE')]
        all_up = all_up[all_up.len.min(axis=1)>30]
        if (dim.endswith('Days') | dim.endswith('Hours') | dim.endswith('Mins')):
            all_up.sort_index(inplace=True)
        
    return all_up.round(4)


def custom_impact_pivot_admin(df, original_interval,previous,current):
    '''Inputs: Data, and two lists of monthly periods - in 'YYY-MM' format- 
    and signals is a [LIST] of variables for which we want to examine impact/contribution
    Also, a [LIST] of OsBuilds to Filter by'''
    #periods = [pd.Period(previous),pd.Period(current)] #Specifying the intervals for comparison
    
    if 'Custom_Period' in df.columns:
        df.drop('Custom_Period',axis=1,inplace=True)
    
    if original_interval.endswith('_year'):
        previous_periods = [pd.Period(x) for x in previous]
        current_periods = [pd.Period(y) for y in current]
    else:
        previous_periods = previous.copy()
        current_periods = current.copy()

    previous = ''.join(previous)
    current = ''.join(current)
    
    for p in previous_periods:
        print(p)
        df.loc[df[original_interval]==p,'Custom_Period'] = previous
    #df.loc[df[time_interval].isin(previous_periods),'Custom_Period'] = previous
    for c in current_periods:
        df.loc[df[original_interval]==c,'Custom_Period'] = current

    df = df[df['Custom_Period'].isin([previous,current])] # filter for faster computation
    
    print(df.Custom_Period.value_counts())
    
    all_up_table = make_admin_pivot_new(df, 'Custom_Period', previous, current)
    all_up_table.index = ['AllUp']
    all_up_table.to_csv('Admin_Impact_Custom'+d+'.csv')
        
    slices = ['Feedback_Source',
              #Tenant Fields
             'Tenant_MSSalesAreaName','Tenant_CommunicationLanguage',#'Tenant_TenantType',
             'Tenant_Type','Tenant_AccountType',
             'Tenant_HasEducation','Tenant_HasCharity','Tenant_HasGovernment',
             'Tenant_HasExchange','Tenant_HasLync','Tenant_HasSharePoint',
             #'Tenant_HasProPlus','Tenant_HasPaid',
             'Tenant_HasYammer','Tenant_HasSubscription',
             'Tenant_HasProject','Tenant_HasVisio','Tenant_HasTrial','Tenant_HasNonTrial',
             'Tenant_IsConcierge',#'Tenant_ConciergeInfoIsManualAdmittance',
             'Tenant_DirectoryExtensionsSyncEnabled','Tenant_DirSyncEnabled',
             'Tenant_PasswordSyncEnabled','Tenant_PasswordWriteBackEnabled',
             'Tenant_IsDonMT','Tenant_IsViral','Tenant_IsTest',
             'Tenant_IsFastTrackTenant',#'Tenant_IsQuickStart',
             'Tenant_IsRestrictRmsViralSignUp','Tenant_IsMSODSDeleted',
             'Tenant_CustomerSegmentGroup',  #'Tenant_CPTenant_ChannelName',
             'Tenant_MSSalesSubRegionClusterGroupingName',
             'Tenant_MSSalesSegmentName','Tenant_MSSalesSubSegmentName',
             'Tenant_IsM365', 'Tenant_HasM365PaidSeats','Tenant_HasM365SKUEdu',
             'Tenant_HasM365SKUBusiness','Tenant_HasM365SKUF1',	
             'Tenant_HasM365SKUE3','Tenant_HasM365SKUE5',	
             'Tenant_HasOfficeSKUE1','Tenant_HasOfficeSKUE3',
             'Tenant_HasOfficeSKUE4','Tenant_HasOfficeSKUE5',
             'Tenant_TenantCategory',
             #'Tenant_AssignedPlanCount_Ord','Tenant_TotalGroupCount_Ord',
             'Tenant_TotalSubscriptionsCount_Ord','Tenant_TotalUsers_Ord',
             #'Tenant_Odin_LicensedUsers_Ord','Tenant_Odin_TotalUsers_Ord',
            
              #Admin Fields
             'Admin_UsageSegment','Admin_PortalVisits_15Days_Ord','Admin_PortalVisits_30Days_Ord',
              #LastTicket Fields
             'LastTicket_IsPartner','LastTicket_Modality','LastTicket_PremierOffering',
             'LastTicket_ProgramName','LastTicket_Source','LastTicket_TeamName',
#              'LastTicket_TimeToLastTicketDays','LastTicket_TimeToLastTicketHours','LastTicket_TimeToLastTicketMins',
#              'LastTicket_TimeToResolveDaysOrd','LastTicket_TimeToResolveHoursOrd','LastTicket_TimeToResolveMinsOrd',
              
              #Subscription Fields
              'Subscription_Sum_O365BusinessPremium','Subscription_Sum_O365BusinessEssential','Subscription_Sum_O365Business',
              'Subscription_Sum_O365E1','Subscription_Sum_O365E2','Subscription_Sum_O365E3','Subscription_Sum_O365E4','Subscription_Sum_O365E5',
              'Subscription_Sum_O365A1','Subscription_Sum_O365A3','Subscription_Sum_O365A5','Subscription_Sum_O365F1',
              'Subscription_Sum_WinE3','Subscription_Sum_WinE5','Subscription_Sum_PowerBI','Subscription_Sum_EMSE3','Subscription_Sum_EMSE5',
              'Subscription_Sum_AudioConferencing','Subscription_Sum_CallingPlan','Subscription_Sum_Dynamics','Subscription_Sum_Planner',
              'Subscription_Sum_Intune','Subscription_Sum_ExtraStorage',

              #Verbatim Fields
            #  'Verbatim_Sum_AdminTheme','Verbatim_Sum_SkypeTheme',
            #  'Verbatim_Sum_OtherTheme','Verbatim_Sum_MarketingTheme',
            #  'Verbatim_Sum_MobileTheme','Verbatim_Sum_ThemeNotMapped',
            #  'Verbatim_Sum_OutlookTheme','Verbatim_Sum_OneDriveTheme',
            #  'Verbatim_Sum_OfficeClientTheme','Verbatim_Sum_CustomerSupportTheme',
            #  'Verbatim_Sum_CommerceTheme','Verbatim_Sum_SharepointTheme',
            #  'Verbatim_Sum_ExchangeTheme','Verbatim_Sum_TeamsTheme','Verbatim_Sum_YammerTheme'
            ]
    
    df_empty = pd.DataFrame({' ' : []})
    df_empty.to_csv('Admin_Impact_Custom'+d+'.csv',mode='a')
    for var in slices:
        print(var)
        table = make_admin_pivot_new(df, 'Custom_Period', previous, current, dim=var)
        #table = table.sort_index()
        pd.DataFrame(data=[var]).to_csv('Admin_Impact_Custom'+d+'.csv',mode='a',index=False,header=False)
        table.to_csv('Admin_Impact_Custom'+d+'.csv',mode='a', header=False)
        df_empty.to_csv('Admin_Impact_Custom'+d+'.csv',mode='a')
    
    all_up_table2 =  make_admin_pivot_new(df, 'Custom_Period', previous, current,dim=slices[0])
    
    #print(df['Tenant_Odin_LicensedUsers_Ord'].value_counts())
    #print(df['Tenant_Odin_TotalUsers_Ord'].value_counts())

    sources = df.Feedback_Source.dropna().unique()
    if len(sources)>1:
        for source in df.Feedback_Source.dropna().unique():
            all_up_table.to_csv('Admin_Impact_'+source+d+'.csv')
            all_up_table2[all_up_table2.index==source].to_csv('Admin_Impact_'+source+d+'.csv',mode='a',header=False)
            data = df[df.Feedback_Source==source]
            table = make_admin_pivot_new(data, 'Custom_Period',previous, current)
            table.index = ['AllUp'+source]
            table.to_csv('Admin_Impact_'+source+d+'.csv',mode='a',header=False)
            df_empty.to_csv('Admin_Impact_'+source+d+'.csv',mode='a')
            for var in slices[1:]:
                table = make_admin_pivot_new(data, 'Custom_Period', previous, current, dim=var)
                pd.DataFrame(data=[var]).to_csv('Admin_Impact_'+source+d+'.csv',mode='a',index=False,header=False)
                table.to_csv('Admin_Impact_'+source+d+'.csv',header=False,mode='a')
                df_empty.to_csv('Admin_Impact_'+source+d+'.csv',mode='a')

# prev = [str(df.Feed_week.max()-1)]
# curr = [str(df.Feed_week.max())]

#custom_impact_pivot_admin(df,'month_year',
            # ['2019-05'],
            # ['2019-06'])

# import xlwt
# import glob
# import pyexcel as p

# num_check = set('0123456789.')
# letter_check = set('qwertyuiopasdfghjklzxcvbnm-+()[]{}/')
# wb = xlwt.Workbook()
# for filename in glob.glob("Admin*"+d+".csv"):
#     (f_path, f_name) = os.path.split(filename)
#     (f_short_name, f_extension) = os.path.splitext(f_name)
#     print(f_short_name)
#     ws = wb.add_sheet(f_short_name[:31])
#     fileReader = csv.reader(open(filename, 'rt'))
#     for rowx, row in enumerate(fileReader):
#         for colx, value in enumerate(row):
#             if (any((n in value) for n in num_check)) & (any((c in value.lower()) for c in letter_check) is False):
#                 num_value = float(value)
#                 ws.write(rowx, colx, num_value)
#             elif value.strip().startswith('-'):
#                 num_value = float(value)
#                 ws.write(rowx, colx, num_value)
#             else:
#                 ws.write(rowx, colx, value)
                
                    
# wb.save("AdminImpactReports"+d+".xls")
# p.save_book_as(file_name='AdminImpactReports'+d+'.xls',
#                dest_file_name='AdminImpactReports'+d+'.xlsx')


#df[df.Feedback_Verbatim.notnull()].to_csv('adminVerbatims'+d+'.csv')
