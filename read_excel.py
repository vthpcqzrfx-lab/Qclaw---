# -*- coding: utf-8 -*-
import pandas as pd
import json
import sys

files = [
    r'C:\Users\zhaoyang26\Desktop\11-12月素养信息流数据\信息流分析-头条&广点通-1月.xlsx',
    r'C:\Users\zhaoyang26\Desktop\11-12月素养信息流数据\小学信息流用户信息.xlsx',
    r'C:\Users\zhaoyang26\Desktop\11-12月素养信息流数据\线索明细数据 0205使用.xlsx'
]

results = []
for f in files:
    try:
        xl = pd.ExcelFile(f)
        info = {'file': f.split('\\')[-1], 'sheets': xl.sheet_names}
        df = pd.read_excel(f, nrows=5)
        info['columns'] = [str(c) for c in list(df.columns)[:10]]
        results.append(info)
    except Exception as e:
        results.append({'file': f.split('\\')[-1], 'error': str(e)})

print(json.dumps(results, ensure_ascii=False, indent=2))
