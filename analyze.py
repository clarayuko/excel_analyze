# -*- coding:UTF-8 -*-


import pandas as pd
import os


path="files/"
files = os.listdir(path)
print(files)
for file in files:
    res={}
    res1={}
    data = pd.read_excel("files/"+file,sheet_name='漏洞信息',usecols="D,S")
    df_vuln=data['漏洞名称'].values.tolist()
    df_way=data["解决办法"].values.tolist()
    result = []
    for i in range(len(df_vuln)):
        if df_vuln[i] not in res:
            res[df_vuln[i]]=os.path.splitext(file)[0]
            res1[df_vuln[i]]=df_way[i]
        else:
            res[df_vuln[i]].append(os.path.splitext(file)[0])
            #res1[df_vuln[i]]=df_way[i]
#print(res)
#print(res1)

ans1=[]
for key,value in res.items():
    ans={}
    if key in res1:
        ans["漏洞名称"]=key
        ans["相关IP"]= value
        ans["解决办法"]=res1[key]
    ans1.append(ans)
#print(ans1)

pf = pd.DataFrame(ans1)
order = ['漏洞名称','相关IP','解决办法']
pf=pf[order]
file_path = pd.ExcelWriter('text.xlsx')
pf.fillna(' ', inplace=True)
pf.to_excel(file_path, encoding='utf-8', index=False,sheet_name="sheet1")
file_path.save()