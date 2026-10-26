from pathlib import Path
p=Path('app.js');s=p.read_text(encoding='utf-8-sig')
s=s.replace('OA / 集团模板尚未提供，当前为通用 CSV。','已收到 OA V1.2 模板；此处为通用台账 CSV，OA 分册映射见下方。集团人资模板仍待确认。')
s=s.replace('<h3>选择导出字段</h3>${Object.entries(fields)','<h3>选择导出字段</h3><div class="export-fields">${Object.entries(fields)')
s=s.replace('<label style="display:inline-block;margin:10px"><input class="expfield"','<label class="export-option"><input class="expfield"')
s=s.replace('''${t}</label>`).join('')}<p><button onclick="download()">''','''${t}</label>`).join('')}</div><p><button onclick="download()">''')
p.write_text(s,encoding='utf-8')
