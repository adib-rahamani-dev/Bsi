"""Read the source without modifying it. All cell values remain strings for exact integers."""
import sys, json, hashlib
from pathlib import Path
import openpyxl

source = Path(sys.argv[1])
wb = openpyxl.load_workbook(source, data_only=False)
cached = openpyxl.load_workbook(source, data_only=True)
result = {'name': source.name, 'sha256': hashlib.sha256(source.read_bytes()).hexdigest(), 'sheets': []}
for sheet in wb:
    rows = []
    for row in sheet:
        cells = []
        for cell in row:
            value = cached[sheet.title][cell.coordinate].value if cell.data_type == 'f' else cell.value
            cells.append({'v': '' if value is None else str(value), 'formula': cell.value if cell.data_type == 'f' else None, 'error': cell.data_type == 'e', 'missingCache': cell.data_type == 'f' and value is None})
        rows.append(cells)
    result['sheets'].append({'name': sheet.title, 'rows': rows})
Path('private/workbook.json').write_text(json.dumps(result, ensure_ascii=False), encoding='utf8')
print(json.dumps({'sheets': len(result['sheets']), 'rows': sum(len(s['rows']) for s in result['sheets'])}))
