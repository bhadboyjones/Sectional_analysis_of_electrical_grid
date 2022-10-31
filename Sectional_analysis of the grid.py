import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# Unpacking the data
df = pd.read_excel('system.xlsx', sheet_name='system', index_col='Unnamed: 0')
pd.set_option('display.max_columns', 15)
df = df.iloc[:4, :12]

# Create new columns
df['Buses Above'] = ''
df['Buses Below'] = ''

# How to enter input into cells in a df
df.columns
df.columns[:-2] # shows all columns except last two

# To add the str: B5+ to location below
df.loc['L2', 'Buses Below'] = df.loc['L2', 'Buses Below'] + 'B5+'

# Another way
df.loc['L1', 'Buses Above'] += 'B3+'

df['Buses Above'] = ''
df['Buses Below'] = ''

# Loop through all columns except last two
for k in df.columns[:-2]:
    for i in df.index:
        if df.loc[i, k] == 'Above':
            df.loc[i, 'Buses Above'] += k + '+'
        else:
            df.loc[i, 'Buses Below'] += k + '+'

# Removing plus sign at the end of last 2 columns
df['Buses Above'] = df['Buses Above'].str[0:-1]
df['Buses Below'] = df['Buses Below'].str[0:-1]

# Analysing
#print(df[['b1', 'b2']])
df2 = df[['Buses Above', 'Buses Below']]
df_buses = pd.DataFrame({'Buses': ['bus_1', 'bus_2', 'bus_3', 'bus_4', 'bus_5', 'bus_6', 'bus_7', 'bus_8', 'bus_9',
                                   'bus_10', 'bus_11', 'bus_12']})

df_buses['Demand (MW)'] = [0, 3, 3, 1, 6, 4, 3, 7, 2, 4, 6, 7]
print(df_buses)
df2['total demand Above'] = ''
df2['total demand Below'] = ''

# Preventing warnings
df2_copy = df2.copy()
df2_copy = df2_copy[['Buses Above', 'total demand Above', 'Buses Below', 'total demand Below']]
df2_copy
print('---------------')

'''
# Assigning
globals()['test'] = 15
alpha = 'new'
globals()['test_%s' %alpha] = 12
print(globals())
print(range(1, 12))
'''

# Assign values for b1, b2, b3 and so on based on df_buses
for i in range(1, len(df_buses)+1):
    globals()['b%s' %i] = df_buses['Demand (MW)'][i-1]
b1, b2, b3


print(eval(df2_copy.iloc[0, 0]))
print(eval(df2_copy.iloc[3, 0]))

for i in range(0, len(df2_copy)):
    df2_copy.iloc[i, 1] = eval(df2_copy.iloc[i, 0])
    df2_copy.iloc[i, 3] = eval(df2_copy.iloc[i, 2])

print(df2_copy)
with pd.ExcelWriter('above_below.xlsx') as writer:
    df2.to_excel(writer, sheet_name='topology')

