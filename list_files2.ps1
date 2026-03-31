Get-ChildItem -Path "C:\Users\zhaoyang26\Desktop" -Include "*.xlsx","*.xls","*.docx" -Recurse | Select-Object Name, Length
