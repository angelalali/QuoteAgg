import pandas as pd
import os
import time
from pyexcelerate import Workbook

start = time.time()
rootDir = r'C:/Users/yisli/Documents/landlordlady/Tesla/ad hoc projects/part price breakdown/output/all/7-27/'

data_X = pd.read_excel(os.path.join(rootDir,'X.xlsx'))
data_X['Program'] = 'X'
data_S = pd.read_excel(os.path.join(rootDir,'S.xlsx'))
data_S['Program'] = 'S'
data_3 = pd.read_excel(os.path.join(rootDir,'3.xlsx'))
data_3['Program'] = '3'

data_all = data_X.append(data_S.append(data_3))
data_all.reset_index(inplace=True)

# print('class of qutoedate is: ', type(data_all['QuoteDate']))
# print('and the data looks like: ', data_all['QuoteDate'][10])

# data_all.loc[data_all['QuoteDate'].notnull(), 'QuoteDate'] = (pd.to_datetime('QuoteDate') - data_all['QuoteDate']).dt.days
# print('im going to start printing dates!')
# print(data_all['QuoteDate'][0:10])

merged_file = os.path.join(rootDir, 'price breakdown all programs.xlsx')
print(merged_file)
if os.path.exists(merged_file):
    os.remove(merged_file)

print('time elapsed: ', time.time() - start)

data_all.to_excel(merged_file, index=False)
# def df_to_excel(df, path, sheet_name='Sheet 1'):
#     data = [df.columns.tolist(), ] + df.values.tolist()
#     wb = Workbook()
#     wb.new_sheet(sheet_name, data=data)
#     wb.save(path)
#     wb.close()
#
# df_to_excel(df=data_all, path=merged_file)

print('job finished!')
print('time elapsed: ', time.time() - start)

