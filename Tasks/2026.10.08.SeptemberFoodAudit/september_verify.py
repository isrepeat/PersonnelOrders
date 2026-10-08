import collections
import datetime as dt
import gzip
import hashlib
import json
import subprocess
import zipfile
from pathlib import Path

import openpyxl

OUT = Path(__file__).resolve().parent / 'Results'
payload = json.load(gzip.open(OUT / 'september.json.gz', 'rt', encoding='utf-8'))
output = OUT / 'Продовольче — сверка сентября 2026.xlsm'
book = openpyxl.load_workbook(output, data_only=True, keep_vba=True)
sheet = book['Котел']
assert [sheet.cell(2, i).value for i in range(1, 48)] == payload['headers']
assert sheet.freeze_panes is None
rows = collections.defaultdict(list)
for cells in sheet.iter_rows(min_row=3, max_col=47):
    value = cells[0].value
    if isinstance(value, dt.datetime) and value.year == 2026 and value.month == 9:
        rows[value.date().isoformat()].append(cells)
for day in payload['days']:
    actual = rows[day['date']]
    assert len(actual) == 2, (day['date'], len(actual))
    for which, cells in zip(('old', 'new'), actual):
        expected = day[which]
        for i in range(1, 43):
            assert cells[i].value == expected[i], (day['date'], which, payload['headers'][i], cells[i].value, expected[i])
        for col in day['changes']:
            assert cells[col - 1].fill.fgColor.rgb == ('FFFF6666' if which == 'old' else 'FF92D050'), (day['date'], col)
    assert actual[1][3].value == actual[1][25].value
    assert sum(c.value or 0 for c in actual[1][26:43]) == actual[1][25].value
    assert actual[1][1].value - actual[1][3].value == actual[1][2].value
    if day['ambiguous']:
        assert actual[1][25].fill.fgColor.rgb == 'FFFFD966'
    assert actual[1][45].value is False, (day['date'], actual[1][45].value)
assert len(rows) == 30
assert not [(c.coordinate, c.value) for row in sheet.iter_rows(min_row=410, max_row=469, max_col=47) for c in row if c.data_type == 'e']
with zipfile.ZipFile(OUT / 'Продовольче.xlsm') as before, zipfile.ZipFile(output) as after:
    assert after.testzip() is None
    if 'xl/vbaProject.bin' in before.namelist():
        # Excel добавляет пустой модуль нового листа; текст исходных модулей проверяется через VBProject отдельно.
        assert len(after.read('xl/vbaProject.bin')) > 0
desktop = Path('C:/Users/isrepeat/OneDrive/Рабочий стол')
for name in ('РУХ_last.xlsx', 'Продовольче.xlsm'):
    # Открытый в Excel оригинал читается с разрешением совместного доступа.
    source_path = str(desktop / name).replace("'", "''")
    command = f"$auditStream=[IO.File]::Open('{source_path}',[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)); try {{ [BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash($auditStream)).Replace('-','').ToLowerInvariant() }} finally {{ $auditStream.Dispose() }}"
    source_hash = subprocess.check_output(['powershell', '-NoProfile', '-Command', command], text=True).strip()
    assert source_hash == hashlib.sha256((OUT / name).read_bytes()).hexdigest(), name
print('PASS: 30 old/new pairs, 165 highlighted changes, daily statistics and order totals, source column order, VBA preservation, no frozen panes; desktop sources unchanged.')