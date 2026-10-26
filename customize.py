import pathlib,json,xlrd,openpyxl
root=pathlib.Path('OA个人信息填报'); groups=[]
for p in sorted(root.rglob('*.xls')):
 if not p.name[0].isdigit() or int(''.join(c for c in p.name.split('员工')[0] if c.isdigit()) or '99')>13: continue
 b=xlrd.open_workbook(str(p));s=next((s for s in b.sheets() if s.name!='填写指南'),b.sheet_by_index(0));fs=[]
 for c in range(s.ncols):
  title=str(s.cell_value(1,c)).strip(); code=str(s.cell_value(2,c)).strip()
  if title and code and 'C_OID' not in code and 'C_OPERATOR' not in code and 'C_OPERATE_TIME' not in code and 'C_EMPLOYEE_ID' not in code:
   fs.append({'title':title,'code':code,'required':'*' in title or '＊' in title})
 groups.append({'name':p.stem.strip(),'fields':fs,'multi':p.name[0] not in ['1','2','7','9'] or p.name.startswith(('10','11','12'))})
groups.sort(key=lambda g:int(''.join(c for c in g['name'].split('员工')[0] if c.isdigit())))
pathlib.Path('schema.js').write_text('const SCHEMA='+json.dumps(groups,ensure_ascii=False)+';',encoding='utf-8')
for p in root.rglob('*.xlsx'):
 if p.name.startswith('~$'):continue
 w=openpyxl.load_workbook(p,data_only=True,read_only=True)
 print(p.name,[(s.title,s.max_row,s.max_column) for s in w]);print('headers',list(w.active.values)[0])
md=['# OA V1.2 字段映射\n','来源：用户提供的原始模板。星号按模板保留；系统关联字段不展示为员工输入项。\n']
for g in groups:
 md+=['## '+g['name'],'|界面字段|模板字段标识|必填标记|','|---|---|---|']
 md += ['|'+f['title'].replace('|','/')+'|'+f['code']+'|'+('是' if f['required'] else '否')+'|' for f in g['fields']]
pathlib.Path('字段映射.md').write_text('\n'.join(md),encoding='utf-8')
print('groups',len(groups))
