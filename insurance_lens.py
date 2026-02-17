import numpy as np
import pandas as pd 

#########################Instructions#############################
#Change excel_path, sheet_name, usecols to match excel formatting#
##################################################################

df = pd.read_csv(r"C:\\Users\\matthewsm\\OneDrive - Enstargroup\\2026 Work\\02. Feb\\three_lens_tst\\One-Year Reserve Risk by Class.csv")
df = df.dropna(how = 'all', axis = 1) #Drop all N/A cols


reserves = pd.read_excel(r"C:\\Users\\matthewsm\\OneDrive - Enstargroup\\2026 Work\\02. Feb\\three_lens_tst\\net_reserves.xlsx")
reserves.columns = ["class", "reserve_type", "year", "amount"]
reserves = reserves.pivot_table(index="class", columns="reserve_type", values="amount", aggfunc="sum").fillna(0)
reserves['net'] = reserves['Gross Reserves'] + reserves['RI Reserves']

net_res = reserves['net'].to_frame().T 
net_res = net_res.rename(columns = {'Total Insurance': 'total'})
net_res['qbe2'] = net_res['qbe_atom23_lloyds'] + net_res['qbe_atom23_other']
net_res.insert(1, 'qbe2', net_res.pop('qbe2'))
net_res = net_res.drop(columns=['qbe_atom23_lloyds', 'qbe_atom23_other'])




'''
#Section 1: Finding Net Reserves
'''

'''
#Section 1: Finding 95th Percentile Losses
'''
no_sims = len(df)
sims_selected = 101

#upper_bound = int(no_sims*0.05 + sims_selected//2)
#lower_bound = int(no_sims*0.05 - sims_selected//2)
start = 9949
end = 10050

#Combining QBE2 into one class
df['qbe2'] = df['qbe_atom23_lloyds'] + df['qbe_atom23_other']
df.insert(1, 'qbe2', df.pop('qbe2'))
df = df.drop(columns=['qbe_atom23_lloyds', 'qbe_atom23_other'])

net_res = net_res.reindex(columns=df.columns, fill_value=0)
print(net_res)


df['total'] = df.iloc[:,1:].sum(axis = 1) # Creating total col excl. sim number
df = df.sort_values('total') #Sorting

#Creating Undiversified 1:20
undiversified = df.iloc[:,1:].quantile(0.05)
#Taking Diversified 1:20
diversified = df.iloc[start:end, 1:].mean(axis = 0, skipna = True)
#Combining into single DF
var_95 = pd.concat([undiversified, diversified], axis = 1).T

print(var_95)

#var_95.to_excel('test_df.xlsx')
