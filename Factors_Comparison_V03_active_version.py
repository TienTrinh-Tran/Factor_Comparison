# -*- coding: utf-8 -*-
"""
Created on Wed Aug 23 10:17:20 2017

@author: ttrinhtran

9/7/17: 
    Version 1.0 using Carrier to match
    Version 2.0: will match by Company Code & NAIC and include enhancements, bug fixes if found during testing trial

    Requirements:
        csv files generated by IQ need to be saved in data\results\st folder
        the files from THIS year build are to be moved to subfolder named New
        the files from LAST year build are to be moved to subfolder named Old        
"""

import numpy as np
import pandas as pd
import win32com.client as win32

states = {
        'AL': ['Alabama'],
        'AK': ['Alaska'],
        'AZ': ['Arizona'],
        'AR': ['Arkansas'],
        'CA': ['California'],
        'CO': ['Colorado'],
        'CT': ['Connecticut'],
        'DE': ['Delaware'],
        'DC': ['Washington DC'],
        'FL': ['Florida'],
        'GA': ['Georgia'],
        'HI': ['Hawaii'],
        'ID': ['Idaho'],
        'IL': ['Illinois'],
        'IN': ['Indiana'],
        'IA': ['Iowa'],
        'KS': ['Kansas'],
        'KY': ['Kentucky'],
        'LA': ['Louisiana'],
        'ME': ['Maine'],
        'MD': ['Maryland'],
        'MA': ['Massachusetts'],
        'MI': ['Michigan'],
        'MN': ['Minnesota'],
        'MS': ['Mississippi'],
        'MO': ['Missouri'],
        'MT': ['Montana'],
        'NE': ['Nebraska'],
        'NV': ['Nevada'],
        'NH': ['New Hampshire'],
        'NJ': ['New Jersey'],
        'NM': ['New Mexico'],
        'NY': ['New York'],
        'NC': ['North Carolina'],
        'ND': ['North Dakota'],
        'OH': ['Ohio'],
        'OK': ['Oklahoma'],
        'OR': ['Oregon'],
        'PA': ['Pennsylvania'],
        'RI': ['Rhode Island'],
        'SC': ['South Carolina'],
        'SD': ['South Dakota'],
        'TN': ['Tennessee'],
        'TX': ['Texas'],
        'UT': ['Utah'],
        'VT': ['Vermont'],
        'VA': ['Virginia'],
        'WA': ['Washington'],
        'WV': ['West Virginia'],
        'WI': ['Wisconsin'],
        'WY': ['Wyoming']
        }

"""_____GET inputs from user_____"""

state = input('Your state, use abbreviation e.g. wa: ' ).upper()
while True:
    if state in states:
        break
    else:
        print('We do not have business in the entered state, please re-enter')
        state = input('Your favorite state, use abbreviation e.g. wa: ' ).upper()
    
st = states.get(state)[0]

line = input('Auto or Home: ' ).capitalize()
while True:
    if line == 'Auto' or line == 'Home':
        break
    else:
        print('Incorrect line of business, please re-enter')
        line = input('Auto or Home: ' ).capitalize()

Old = input('Old Build Date: ')
New = input('New Build Date: ')
                   
sit_num_input = input("Type situation number(s), if multiple, separate Sit# by space: ")      
sit_num = list(sit_num_input.split())

"""_____Companies_____"""

try:
    Old_Co = pd.read_csv(r'T:\ACT_PID\Pcm_brm\%s\InsurQuote\data\results\%s\%s\Companies.csv' %(line, st, Old))
    New_Co = pd.read_csv(r'T:\ACT_PID\Pcm_brm\%s\InsurQuote\data\results\%s\%s\Companies.csv' %(line, st, New))
except OSError:
    print('Please check if Companies.csv file(s) is/are in the folder(s)')

Old_Co.insert(0, 'Version', 'Old')
New_Co.insert(0, 'Version', 'New')
    
Companies = pd.concat([Old_Co, New_Co])
Col_Names = [col for col in New_Co.columns] + [col for col in Companies.columns if col not in [col for col in New_Co.columns]]
Companies = Companies[Col_Names]

Companies = Companies.sort_values(['Carrier Name', 'Version'], ascending=False)

"""_____Symbols_____"""

try:
    Old_Sym = pd.read_csv(r'T:\ACT_PID\Pcm_brm\%s\InsurQuote\data\results\%s\%s\Symbols.csv' %(line, st, Old))
    New_Sym = pd.read_csv(r'T:\ACT_PID\Pcm_brm\%s\InsurQuote\data\results\%s\%s\Symbols.csv' %(line, st, New))
except OSError:
    print('Please check if Symbols.csv file(s) is/are in the folder(s)')

Old_Sym.insert(0, 'Version', 'Old')
New_Sym.insert(0, 'Version', 'New')
    
Symbols = pd.concat([Old_Sym, New_Sym])
Sym_Col_Names = [col for col in New_Sym.columns] + [col for col in Symbols.columns if col not in [col for col in New_Sym.columns]]
Symbols = Symbols[Sym_Col_Names]

Symbols = Symbols.sort_values(['Version', 'Policy Key', 'Vehicle'], ascending=[0, 1, 1])

"""_____Premiums_____"""

try:
    Old_Prem00 = pd.read_csv(r'T:\ACT_PID\Pcm_brm\%s\InsurQuote\data\results\%s\%s\Premiums.csv' %(line, st, Old))
    New_Prem00 = pd.read_csv(r'T:\ACT_PID\Pcm_brm\%s\InsurQuote\data\results\%s\%s\Premiums.csv' %(line, st, New))
except OSError:
    print('Please check if Premiums.csv file(s) is/are in the folder(s)')

Old_Prem00.drop(Old_Prem00.columns[4:9], axis=1, inplace=True)
New_Prem00.drop(New_Prem00.columns[4:9], axis=1, inplace=True)

sit_num_conversion = []
sit_num_conversion = ['SIT0' + x if len(x) == 1 else 'SIT' + x for x in sit_num]

#Old_Prem0 = []
#New_Prem0 = []
#for i in sit_num_conversion:
#    Old_Temp = Old_Prem00.loc[Old_Prem00['Policy No'] == i]
#    New_Temp = New_Prem00.loc[New_Prem00['Policy No'] == i]
#    Old_Prem0 = pd.concat(Old_Prem0, Old_Temp)
#    New_Prem0 = pd.concat(New_Prem0, New_Temp)

Old_Prem0 = Old_Prem00.loc[Old_Prem00['Policy No'].isin(sit_num_conversion)]
New_Prem0 = New_Prem00.loc[New_Prem00['Policy No'].isin(sit_num_conversion)]
  
Old_Prem = pd.merge(Old_Co.loc[:, ['Company Key', 'Company Name', 'Carrier Name']], Old_Prem0, on=['Company Key'], how='left')
New_Prem = pd.merge(New_Co.loc[:, ['Company Key', 'Company Name', 'Carrier Name']], New_Prem0, on=['Company Key'], how='left')

Old_Prem.insert(0, 'Version', 'Old')
New_Prem.insert(0, 'Version', 'New')

Old_Prem = Old_Prem.sort_values(['Policy Key', 'Carrier Name', 'Vehicle', 'Version'], ascending=[1, 1, 1, 0])
New_Prem = New_Prem.sort_values(['Policy Key', 'Carrier Name', 'Vehicle', 'Version'], ascending=[1, 1, 1, 0])    

if list(Old_Prem) != list(New_Prem):
    Premiums = pd.concat([Old_Prem, New_Prem])
    Prem_Col_Names = [col for col in New_Prem.columns] + [col for col in Premiums.columns if col not in [col for col in New_Prem.columns]]
    Premiums = Premiums[Prem_Col_Names]    
    Premiums = Premiums.sort_values(['Policy Key', 'Carrier Name', 'Vehicle', 'Version'], ascending=[1, 1, 1, 0])
#pd.options.display.float_format = '{:.2f}%'.format
else:
    Change_Prem = New_Prem.iloc[:,7:].div(Old_Prem.iloc[:,7:])-1
    #Change_Prem.style.format('{:.2%}'.format)
    Premiums = pd.concat([Old_Prem, New_Prem, Change_Prem])
    Prem_Col_Names = [col for col in New_Prem.columns] + [col for col in Premiums.columns if col not in [col for col in New_Prem.columns]]
    #Premiums = Premiums[Prem_Col_Names]
    
    Premiums['Index'] = Premiums.index
    Prem_Col_Names.insert(0,'Index')
    #Premiums = Premiums[Premiums['Index'],Prem_Col_Names]
    Premiums = Premiums.sort_values(['Index', 'Policy Key', 'Carrier Name', 'Vehicle', 'Version'], ascending=[1, 1, 1, 1, 0])
    Premiums = Premiums[Prem_Col_Names]
    Premiums['Version'].fillna('Change', inplace=True)
    Premiums['Carrier Name'].ffill( inplace=True)
    Premiums['Vehicle'].ffill( inplace=True)
    Premiums['Policy No'].ffill( inplace=True)
    Premiums = Premiums.sort_values(['Policy No', 'Index', 'Carrier Name', 'Vehicle', 'Version'], ascending=[1, 1, 1, 1, 0])

#Premiums.loc[:,'Version'] = Premiums.loc[:,'Version'].fillna(method='ffill', inplace=True)
#for i in range (2, len(Premiums.index)+1, 3):
#    Premiums.iloc[i,1] = Premiums.iloc[i,1].replace('New', 'Change')
#Premiums 
#
#data1 = {"a":[1.,3.,5.,2.],
#         "b":[4.,2,3.,7.],
#         "c":[5.,45.,36,34]}
#data2 = {"a":[4.],
#         "b":[2.],
#         "c":[11.]}
#data3 = {"a":[4.,0,2,2],
#         "b":[2.,2,2,2],
#         "c":[2,2,2,1]}
#
#df1 = pd.DataFrame(data1)
#df2 = pd.DataFrame(data2) 
#df3 = pd.DataFrame(data3) 
#
#dfx = df1.div(df2.ix[0],axis='columns')
#dfx1 = (df1/df2-1).values[0,:]
#dfx2 = df2.div(df1.ix[0],axis='columns').subtract(df1.ix[0],axis='columns')
#dfx = df1.iloc[1:,:].div(df3.iloc[1:,:])
#
#try:
#    Change_Prem = Old_Prem.div(New_Prem)-1
#except TypeError:
#    pass
#def pct_change(Premiums):
#    try:
#        Premiums.iloc[2] = Premiums.iloc[1] - Premiums.iloc[0]
#        return Premiums
#    except TypeError:
#        pass

#Premiums.drop(Premiums.columns[[0,1,2,4]], axis=1, inplace=True)
#pct_change = Premiums.groupby([Premiums['Carrier Name'], Premiums['Policy No'],Premiums['Vehicle']]).pct_change(periods=1)
#means = Premiums.groupby([Premiums['Carrier Name'], Premiums['Policy No'],Premiums['Vehicle']]).mean()
#
#grouped = Premiums.groupby([Premiums['Carrier Name'],Premiums['Vehicle'],Premiums['Policy No']])
#Premiums.pct_change(periods=2)

"""_____Factors_____"""

try:
    Old_factors0 = pd.read_csv(r'T:\ACT_PID\Pcm_brm\%s\InsurQuote\data\results\%s\%s\Factors.csv' %(line, st, Old))
    New_factors0 = pd.read_csv(r'T:\ACT_PID\Pcm_brm\%s\InsurQuote\data\results\%s\%s\Factors.csv' %(line, st, New))
except OSError:
    print('Please check if Factors.csv file(s) is/are in the folder(s)')
    
#Old_factors1 = Old_factors0.sort_values(['Policy Key', 'Company Key', 'Factor Name'], ascending=True)
#New_factors1 = New_factors0.sort_values(['Policy Key', 'Company Key', 'Factor Name'], ascending=True)

#Old_factors2 = pd.merge(Old_factors1, Old_Co.loc[:, ['Company Key', 'Carrier Name', 'Company Name']], on=['Company Key'], how='left')
#New_factors2 = pd.merge(New_factors1, New_Co.loc[:, ['Company Key', 'Carrier Name', 'Company Name']], on=['Company Key'], how='left')


Old_factors2 = pd.merge(Old_factors0, Old_Co.loc[:, ['Company Key', 'Carrier Name', 'Company Name']], on=['Company Key'], how='left')
New_factors2 = pd.merge(New_factors0, New_Co.loc[:, ['Company Key', 'Carrier Name', 'Company Name']], on=['Company Key'], how='left')

#pd.formats.format.header_style = None
#pd.io.formats.format.header_style = None
#pd.core.format.header_style = None
writer = pd.ExcelWriter(r'T:\ACT_PID\Pcm_brm\%s\InsurQuote\data\results\%s\Analysis\%s_%s_vs_%s_Sit_%s.xlsx' %(line, st, state, Old, New, sit_num_input))   
Companies.to_excel(writer,'Companies', index=False)
writer.sheets['Companies'].autofilter(0, 0, 0, len(Companies.columns)-1)
writer.sheets['Companies'].set_column('C:C', 30)
Symbols.to_excel(writer,'Symbols', index=False)
writer.sheets['Symbols'].autofilter(0, 0, 0, len(Symbols.columns)-1)
Premiums.to_excel(writer,'Premiums', index=False)
writer.sheets['Premiums'].autofilter(0, 0, 0, len(Premiums.columns)-1)
#writer.sheets['Premiums'].set_column('E:E', 30)
wb  = writer.book

formatx = wb.add_format({'bold': True, 'font_color': 'red', 'bg_color': 'yellow'})
writer.sheets['Symbols'].freeze_panes('E2')
writer.sheets['Premiums'].freeze_panes('I2')    
writer.sheets['Symbols'].conditional_format('E2:AZ31', {'type':'formula', 'criteria':'=E2<>E32', 'format':formatx})
writer.sheets['Symbols'].conditional_format('E32:AZ61', {'type':'formula', 'criteria':'=E2<>E32', 'format':formatx})    

format2 = wb.add_format({'num_format': '0.0%', 'bold': True, 'color': 'red'})
#format3 = wb.add_format({'num_format': '0.0'})
#format10 = wb.add_format('General')
for i in range(3, len(Premiums.index)+1, 3):
    writer.sheets['Premiums'].set_row(i, None, format2)
writer.sheets['Premiums'].set_column('C:D', None, None, {'level': 1, 'hidden': True})    
writer.sheets['Premiums'].set_column('E:E', 30, None, {'collapsed': True})
#writer.sheets['Premiums'].set_column('A:A', format10)
    
for i in sit_num:
    df_Old = [] 
    df_New = []
    df = []

    if line == 'Home':
        k = 5 - len(i)
        n = 'HO_SIT' + k*'0' + i + '#0'
    else:
        if len(i) == 1:
            if state == 'CA':
                n = 'SIT0' + i + '#0'
            else:
                n = 'SIT0' + i + '#0#0'
        else:
            if state == 'CA':
                n = 'SIT' + i + '#0'
            else:
                n = 'SIT' + i + '#0#0'

    df_Old = Old_factors2.loc[Old_factors2['Policy Key'] == n]
    df_New = New_factors2.loc[New_factors2['Policy Key'] == n]

    if line == 'Home': 
        df = pd.merge(df_Old, df_New, on=['Policy Key', 'Carrier Name', 'Factor Name'], how='outer', suffixes=('_Old', '_New'))
    else:
        df = pd.merge(df_Old, df_New, on=['Policy Key', 'Carrier Name', 'Vehicle', 'Factor Name'], how='outer', suffixes=('_Old', '_New'))

    df['Old'] = pd.to_numeric(df['Factor Value_Old'], errors='coerce')
    df['New'] = pd.to_numeric(df['Factor Value_New'], errors='coerce')
        
    def change(x):
        if np.isnan(x['Old']) and np.isnan(x['New']):
            if x['Factor Value_New'] == x['Factor Value_Old']:
                return ""
            else:
                return "diff."
        elif x['Factor Value_New'] == x['Factor Value_Old']:
            return ""
        elif np.isnan(x['New']) and not np.isnan(x['Old']):
            return "factor no longer applied"
        elif not np.isnan(x['New']) and np.isnan(x['Old']):
            return "new factor"        
        else:
            try:
                return x['New']/x['Old']-1
            except ZeroDivisionError:
                return "check"
                    
    df['New/Old-1'] = df.apply(change, axis=1)
    
    df.Old.fillna(df['Factor Value_Old'], inplace=True)
    df.New.fillna(df['Factor Value_New'], inplace=True)
    
    def change_1(x):
        if x['New/Old-1'] == "factor no longer applied" and x['Old'] < 2:
            try:
                return 1/x['Old']-1
            except ZeroDivisionError:
                pass
        if x['New/Old-1'] == "new factor" and x['New'] < 2:     
            try:
                return x['New']-1
            except ZeroDivisionError:
                pass
    df['Adtl.'] = df.apply(change_1, axis=1)

    if line == 'Home':
        df = df.loc[:,['Policy Key', 'Company Name_Old', 'Company Name_New', 'Carrier Name', 'Factor Name', 'Old', 'New', 'New/Old-1', 'Adtl.']]
    else:
        df = df.loc[:,['Policy Key', 'Company Name_Old', 'Company Name_New', 'Carrier Name', 'Factor Name', 'Vehicle', 'Old', 'New', 'New/Old-1', 'Adtl.']]
    
    df_Old.to_excel(writer, i+'Old', index=False)
    df_New.to_excel(writer, i+'New', index=False)
    df.to_excel(writer,'Sit'+i, index=False)
    
    writer.sheets[i+'Old'].freeze_panes('A2')
    writer.sheets[i+'New'].freeze_panes('A2')    
    writer.sheets['Sit'+i].freeze_panes('E2')
    writer.sheets[i+'Old'].autofilter('A1:F1' if line == 'Home' else 'A1:G1') 
    writer.sheets[i+'New'].autofilter('A1:F1' if line == 'Home' else 'A1:G1') 
    
    writer.sheets['Sit'+i].set_column('A:C', None, None, {'level': 1, 'hidden': True})    
    writer.sheets['Sit'+i].set_column('D:D', None, None, {'collapsed': True}) 
    writer.sheets['Sit'+i].autofilter('A1:I1' if line == 'Home' else 'A1:J1')   

    writer.sheets['Sit'+i].set_column('D:D', 30)
    writer.sheets['Sit'+i].set_column('E:E', 60)

    format0 = wb.add_format({'num_format': '0.0%'})
    format1 = wb.add_format({'bold': True})
    writer.sheets['Sit'+i].set_row(0, None, format1)
    writer.sheets['Sit'+i].set_column('H:H' if line == 'Home' else 'I:I', 22, format0)
    writer.sheets['Sit'+i].set_column('I:I' if line == 'Home' else 'J:J', None, format0)


writer.save()

excel = win32.gencache.EnsureDispatch('Excel.Application')

#excel.ScreenUpdating = False

wb = excel.Workbooks.Open(r'T:\ACT_PID\Pcm_brm\%s\InsurQuote\data\results\%s\Analysis\%s_%s_vs_%s_Sit_%s.xlsx' %(line, st, state, Old, New, sit_num_input))  
wb.Visible = False
wb.ScreenUpdating = False

wb.Worksheets('Companies').Rows('1:1').WrapText = True
wb.Worksheets('Symbols').Rows('1:1').WrapText = True
wb.Worksheets('Premiums').Rows('1:1').WrapText = True
wb.Worksheets('Premiums').Columns('A:A').NumberFormat = "0"
wb.Worksheets('Premiums').Columns('G:G').NumberFormat = "0"

wb.Save()
wb.Close()

print('Complete!')

