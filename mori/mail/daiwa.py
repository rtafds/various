#!/usr/bin/python
# -*- Coding: utf-8 -*-
#%%
import pandas as pd
import os

dirpath = os.path.dirname(os.path.abspath(__file__))    
daiwa_mail = pd.read_csv(dirpath+'/daiwamail.csv', header=4, encoding='cp932')

len_daiwa_mail = daiwa_mail.shape[0]
daiwa_addr = '@daiwa.co.jp'
count = 0
for i in range(len_daiwa_mail):
    mail = daiwa_mail.iloc[i, 11]
    
    if not daiwa_addr in mail:
        count += 1

print(count)
