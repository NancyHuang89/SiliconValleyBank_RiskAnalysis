import pandas as pd
pd.set_option('display.max_rows', 500)
## Extract Data Source
#read the balance sheet and RC-B
BSheet = pd.read_excel('Bulk_RC_2022Q4.xlsx', sheet_name = 'RC')
RCB = pd.read_excel('Bulk_RC_2022Q4.xlsx', sheet_name = 'RCB')
POC = pd.read_excel('Bulk_RC_2022Q4.xlsx', sheet_name = 'POC')
#Index IDRSSD and Bank name
POC_slim = POC[['IDRSSD','Financial Institution Name']] 
POC_slim.columns = ['IDRSSD','Financial_Institution_Name']

#PART A
##Banks that follows FFIEC 031 standard
#Slim version Bal Sheet
BS_slim_A = BSheet[['IDRSSD','RCFD0081','RCFD0071','RCFD2170','RCFDJJ34','RCFD1773','RCFDJA22','RCFD5369','RCFDB528','RCFD3123','RCFDB529','RCFD3545']]
BS_slim_A.columns = ['IDRSSD','NIB_Cash','IB_Cash','TOT_ASSETS','HTM','AFS','Total_Loans','Loans_HFS','Loans_AFS','ALLL','Loans_Net','Trading_Assets']
BS_slim_A = BS_slim_A.fillna(0)
#calulate total securities (HTM + AFS), not including Equities
BS_slim_A['TOT_SEC'] = BS_slim_A[['HTM','AFS']].sum(axis=1)
#Slim version RC-B
RCB_slim_A = RCB[['IDRSSD','RCFD1754','RCFD1771']]
RCB_slim_A.columns = ['IDRSSD','HTM_AmortCost','HTM_FairValue']
RCB_slim_A['unreal_HTM_Loss'] = RCB_slim_A['HTM_FairValue'] - RCB_slim_A['HTM_AmortCost']
#Combine the three sources
Comb = POC_slim.merge(BS_slim_A,how='left',on='IDRSSD').merge(RCB_slim_A,how='left',on='IDRSSD')
Comb['Loss_Tot_Asset'] = Comb['unreal_HTM_Loss']/Comb['TOT_ASSETS']
#Sort the Unrealized Loss to Total Assets Ratio
Consolidated_A = Comb.sort_values(by = 'Loss_Tot_Asset', ascending = True)
Consolidated_A = Consolidated_A[Consolidated_A['TOT_ASSETS']>0]
Consolidated_A['Category'] = 'Domestic_And_Foreign_Offices'

#PART A
##Banks that follows FFIEC 041 and FFIEC 051 standard
#Slim version Bal Sheet
BS_slim_B = BSheet[['IDRSSD','RCON0081','RCON0071','RCON2170','RCONJJ34','RCON1773','RCONJA22','RCON5369','RCONB528','RCON3123','RCONB529','RCON3545']]
BS_slim_B.columns = ['IDRSSD','NIB_Cash','IB_Cash','TOT_ASSETS','HTM','AFS','Total_Loans','Loans_HFS','Loans_AFS','ALLL','Loans_Net','Trading_Assets']
BS_slim_B = BS_slim_B.fillna(0)
#calulate total securities (HTM + AFS), not including Equities
BS_slim_A['TOT_SEC'] = BS_slim_A[['HTM','AFS']].sum(axis=1)
#Slim version RC-B
RCB_slim_B = RCB[['IDRSSD','RCON1754','RCON1771']]
RCB_slim_B.columns = ['IDRSSD','HTM_AmortCost','HTM_FairValue']
RCB_slim_B['unreal_HTM_Loss'] = RCB_slim_B['HTM_FairValue'] - RCB_slim_B['HTM_AmortCost']
#Combine the three sources
Comb = POC_slim.merge(BS_slim_B,how='left',on='IDRSSD').merge(RCB_slim_B,how='left',on='IDRSSD')
Comb['Loss_Tot_Asset'] = Comb['unreal_HTM_Loss']/Comb['TOT_ASSETS']
#Sort the Unrealized Loss to Total Assets Ratio
Consolidated_B = Comb.sort_values(by = 'Loss_Tot_Asset', ascending = True)
Consolidated_B = Consolidated_B[Consolidated_B['TOT_ASSETS']>0]
Consolidated_B['Category'] = 'Domestic_Only'


#Export PART A and PART B
Consolidated_A.to_excel('HTM_Unrealized_Loss_Analysis.xlsx',sheet_name = 'Domestic_And_Foreign_Offices',index=False)Final.sort_values(by = 'Loss_Tot_Asset', ascending = True).head(200).to_csv('HTM_unrealizedloss_asset_ratio_highest200.csv')
with pd.ExcelWriter('HTM_Unrealized_Loss_Analysis.xlsx',mode='a',engine = 'openpyxl') as writer:
    #writer.book = openpyxl.load_workbook('HTM_Unrealized_Loss_Analysis.xlsx')
    Consolidated_B.to_excel(writer, sheet_name='Domestic_Only')