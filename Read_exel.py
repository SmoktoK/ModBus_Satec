import pandas as pd
import numpy as np

# cols = [0, 1, 2]
# read_params = pd.read_excel('satec.xlsx', usecols=cols)


read_params = pd.read_excel('satec.xlsx')

# read_params.head()
# read_params.iloc[1, 1]
# print(read_params.iloc[0, 2])
l = read_params.iloc[1, 2]
print(l)
print(type(l))
print('Количество строк:', read_params.shape[0])






# for i in l:
#     print(i)

# print(read_params.columns.ravel())
# ip = pd.read_excel('satec.xlsx', usecols=2)
# print(ip)