# -*- coding: utf-8 -*-
import pandas as pd

# 读取城市P8等级表
df_city = pd.read_excel(r'C:\Users\zhaoyang26\Desktop\【春二轮&春暑秋】-城市P8等级表.xlsx')
print('===城市P8等级表===')
print('列名:', df_city.columns.tolist())
print('行数:', len(df_city))
print(df_city.head(20))
print('\n===各等级分布===')
# 找到等级列
for col in df_city.columns:
    if '等级' in str(col):
        print(df_city[col].value_counts().sort_index())
        break

print('\n\n===工作簿-素材===')
df_material = pd.read_excel(r'C:\Users\zhaoyang26\Desktop\工作簿-素材.xlsx')
print('列名:', df_material.columns.tolist())
print('行数:', len(df_material))
print(df_material.head(20))
print('\n===任务状态分布===')
for col in df_material.columns:
    if '状态' in str(col):
        print(df_material[col].value_counts())
        break
