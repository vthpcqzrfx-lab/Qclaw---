# -*- coding: utf-8 -*-
import pandas as pd
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 读取城市P8等级表
print('='*60)
print('【城市P8等级表分析】')
print('='*60)
try:
    df_city = pd.read_excel(r'C:\Users\zhaoyang26\Desktop\【春二轮&春暑秋】-城市P8等级表.xlsx')
    print(f'总记录数: {len(df_city)}')
    print(f'列名: {list(df_city.columns)}')
    print('\n前15行数据:')
    print(df_city.head(15).to_string())
    print('\n\n数据概览:')
    print(df_city.info())
except Exception as e:
    print(f'读取错误: {e}')

print('\n' + '='*60)
print('【工作簿-素材分析】')
print('='*60)
try:
    df_material = pd.read_excel(r'C:\Users\zhaoyang26\Desktop\工作簿-素材.xlsx')
    print(f'总记录数: {len(df_material)}')
    print(f'列名: {list(df_material.columns)}')
    print('\n前15行数据:')
    print(df_material.head(15).to_string())
    print('\n\n数据类型:')
    print(df_material.dtypes)
except Exception as e:
    print(f'读取错误: {e}')
