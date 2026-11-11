"""Read the workbook using Python's standard library and refresh data.js."""
from pathlib import Path
import json
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / '붙임1_프로그램별_재원.xlsx'
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(SOURCE) as z:
    strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        for item in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si', NS):
            strings.append(''.join(item.itertext()))
    sheet = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    rows = []
    for row in sheet.findall('s:sheetData/s:row', NS):
        n = int(row.attrib['r'])
        if not 5 <= n <= 36:
            continue
        values = {}
        for cell in row.findall('s:c', NS):
            col = ''.join(c for c in cell.attrib['r'] if c.isalpha())
            value = cell.find('s:v', NS)
            kind = cell.attrib.get('t')
            if kind == 's':
                value = strings[int(value.text)]
            elif kind == 'str':
                value = value.text if value is not None else ''
            elif kind == 'inlineStr':
                value = ''.join(cell.find('s:is', NS).itertext())
            else:
                value = float(value.text) if value is not None and value.text else None
            values[col] = value
        a, b = values['E'], values['F']
        assert isinstance(a, (int, float)) and isinstance(b, (int, float))
        rows.append(dict(id=n, sector=values['A'], account=values['B'], group=values['C'], name=values['D'], a=a, b=b, diff=b-a, rate=(b-a)/a if a else None))
assert len(rows) == 32
(ROOT / 'data.js').write_text('window.BUDGET_DATA = ' + json.dumps(rows, ensure_ascii=False, indent=2) + ';\n', encoding='utf-8')
print(f'Updated {len(rows)} records from {SOURCE.name}')
