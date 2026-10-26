import xlrd,pathlib,json
out=[]
for p in pathlib.Path('OA个人信息填报').rglob('*.xls*'):
 if p.name.startswith('~$'):continue
 try:
  b=xlrd.open_workbook(str(p));item={'file':p.name,'sheets':[]}
  for s in b.sheets():
   item['sheets'].append({'name':s.name,'rows':s.nrows,'columns':s.ncols,'preview':[[str(s.cell_value(r,c)) for c in range(s.ncols)] for r in range(min(s.nrows,5))]})
  out.append(item)
 except Exception as e:out.append({'file':p.name,'error':str(e)})
pathlib.Path('模板结构核对.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(out,ensure_ascii=False))
