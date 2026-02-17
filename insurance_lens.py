import numpy as np
import pandas as pd 

##################################################################
#Section 1: import and adjustment of data                        #
##################################################################

invalid_classes = ['micl_cbre_qs_assumed', 'lloyds_assumed','sise_cbre_qs_assumed','rti_cbre_qs_assumed']

reserve_sim_path = r"C:\\Users\\matthewsm\\OneDrive - Enstargroup\\2026 Work\\02. Feb\\three_lens_tst\\One-Year Reserve Risk by Class.csv"
reserve_volume_path = r"C:\\Users\\matthewsm\\OneDrive - Enstargroup\\2026 Work\\02. Feb\\three_lens_tst\\net_reserves.xlsx"

df = pd.read_csv(reserve_sim_path)
df = df.dropna(how = 'all', axis = 1) #Drop all N/A cols

#Combining QBE2 into one class
df['qbe2'] = df['qbe_atom23_lloyds'] + df['qbe_atom23_other']
df.insert(1, 'qbe2', df.pop('qbe2'))
df = df.drop(columns=['qbe_atom23_lloyds', 'qbe_atom23_other'])
df = df.drop(columns= invalid_classes)
df['total'] = df.iloc[:,1:].sum(axis = 1) # Creating total col excl. sim number

#Taking Gross & RI Reserves to calculate Net Reserves
reserves = pd.read_excel(reserve_volume_path)
reserves.columns = ["class", "reserve_type", "year", "amount"]
reserves = reserves.pivot_table(index="class", columns="reserve_type", values="amount", aggfunc="sum").fillna(0)
reserves['net'] = reserves['Gross Reserves'] - reserves['RI Reserves']
reserves = reserves.drop(index = invalid_classes)
tot_net_res = reserves['net'].sum() - reserves.loc['Total Insurance', 'net']

#Taking just net reserves
net_res = reserves['net'].to_frame().T 
net_res = net_res.rename(columns = {'Total Insurance': 'total'}) # Rename of column to align with df
#Combining QBE2 into one class
net_res['qbe2'] = net_res['qbe_atom23_lloyds'] + net_res['qbe_atom23_other']
net_res.insert(1, 'qbe2', net_res.pop('qbe2'))
net_res = net_res.drop(columns=['qbe_atom23_lloyds', 'qbe_atom23_other'])
net_res = net_res.reindex(columns=df.columns, fill_value=0)

##################################################################
#Section 2: Finding 95th Percentile Loss                         #
##################################################################
start = 9949
end = 10050

df = df.sort_values('total') #Sorting


#Creating Undiversified 1:20
undiversified = df.iloc[:,1:].quantile(0.05)
#Taking Diversified 1:20
diversified = df.iloc[start:end, 1:].mean(axis = 0, skipna = True)
#Combining into single DF
var_95 = pd.concat([undiversified, diversified], axis = 1).T
var_95 = var_95.rename(index={0.05: 'undiversified', 0.00: 'diversified'})


##################################################################
#Section 3: Creating Output Dataframe                            #
##################################################################

combined_df = pd.concat([net_res,var_95], axis = 0, join = 'inner') # Inner join takes intersection and drops cols not present in both ('Sim' in this case)

key_cols = combined_df.loc['net'].nlargest(6).index # Take 6 largest Net Res (5 Classes + Total)
combined_df = combined_df[key_cols]


#Convert to % and $M
output_df = combined_df.copy()
output_df.loc['net']  /= tot_net_res
output_df.loc['diversified']  /= output_df.loc['diversified','total']
output_df.loc['undiversified'] /= 1_000_000


#Transpose and swap order
output_df = output_df.drop(columns = 'total').T
output_df = output_df[['net','diversified','undiversified']] 

#Renaming cols
output_df = output_df.rename(columns={'net': 'Mean Contribution','diversified': '95th contribution to the stress','undiversified': 'Standalone VaR ($M)'})

print(output_df)
#Output
output_df.to_excel('three_lens_table.xlsx')
