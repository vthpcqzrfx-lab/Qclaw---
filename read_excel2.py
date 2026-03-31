import pandas as pd
import os
import sys

files = [
    r'C:\Users\zhaoyang26\Desktop\运营岗-实习转正答辩PPT.xlsx',
    r'C:\Users\zhaoyang26\Desktop\关键岗位.xlsx',
    r'C:\Users\zhaoyang26\Desktop\信息流投放-头条&抖音-2版&3版.xlsx',
    r'C:\Users\zhaoyang26\Desktop\city_level_5_251122.xlsx',
    r'C:\Users\zhaoyang26\Desktop\新建 DOCX 文档.docx',
]

output = []
for f in files:
    fname = os.path.basename(f)
    output.append('=' * 60)
    output.append('FILE: ' + fname)
    output.append('=' * 60)
    try:
        xls = pd.ExcelFile(f)
        output.append('Sheets: ' + str(xls.sheet_names))
        for sheet in xls.sheet_names[:3]:
            df = pd.read_excel(f, sheet_name=sheet, header=None, nrows=22)
            output.append('')
            output.append('--- Sheet: ' + sheet + ' ---')
            # write each row
            for i, row in df.iterrows():
                row_data = []
                for val in row:
                    if pd.isna(val):
                        row_data.append('')
                    else:
                        row_data.append(str(val))
                output.append(' | '.join(row_data[:20]))
    except Exception as e:
        output.append('ERROR: ' + str(e))
    output.append('')

with open(r'C:\Users\zhaoyang26\.qclaw\workspace\excel_summary.txt', 'w', encoding='utf-8') as out:
    out.write('\n'.join(output))
print('Done')
