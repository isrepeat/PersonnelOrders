$ErrorActionPreference = 'Stop'
$outputPath = Join-Path $PSScriptRoot 'Анализ смен статусов v29.xlsx'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Анализ смен статусов v28.xlsx') -Destination $outputPath
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open($outputPath)
    $sheet = $book.Worksheets.Item(1)
    foreach ($cache in $book.SlicerCaches) { $cache.ClearManualFilter() }
    if ($sheet.FilterMode) { $sheet.ShowAllData() }
    # Вставка перед служебными идентификаторами сохраняет их и ссылки счётчиков.
    $sheet.Columns.Item('S').Insert()
    $sheet.Columns.Item('S').Hidden = $false
    $sheet.Columns.Item('T:U').Hidden = $true
    $table = $sheet.ListObjects.Item('FourRowReview')
    $table.Resize($sheet.Range('A10:S1230'))
    $sheet.Range('R10:R1230').Copy($sheet.Range('S10:S1230'))
    $sheet.Range('S10').Value2 = 'Уровень безопасности закрытия'
    $sheet.Columns.Item('S').ColumnWidth = 44
    $counts = @{}
    for ($first = 11; $first -le 1230; $first += 4) {
        $status = [string]$sheet.Cells.Item($first, 16).Value2
        $operation = [string]$sheet.Cells.Item($first, 14).Value2
        $unknown = $false
        $changedOther = $false
        $largeShift = $false
        foreach ($offset in @(2,3)) {
            $row = $first + $offset
            if ($sheet.Cells.Item($row, 4).Text -eq 'Не определено') { $unknown = $true }
            foreach ($column in @(4,5,6,9,12,13)) {
                if ([string]$sheet.Cells.Item($row,$column).Value2 -ne [string]$sheet.Cells.Item(($row-2),$column).Value2) { $changedOther = $true }
            }
            foreach ($column in @(8,10)) {
                $before = $sheet.Cells.Item(($row-2),$column).Value2
                $after = $sheet.Cells.Item($row,$column).Value2
                if ($before -and -not $after) { $unknown = $true }
                if ($before -is [double] -and $after -is [double] -and [Math]::Abs($after-$before) -gt 1) { $largeShift = $true }
            }
        }
        if ($unknown -or $status -like '*Требует проверки*' -or $operation -like '*Исправление не подтверждено*') {
            $label = '5 — Ручная проверка: данных недостаточно'
        } elseif ($status -like '*Расхождение с приказом*') {
            $label = '3 — Расхождение с приказом: приоритет РУХ'
        } elseif ($status -like '*Несоответствие*' -or $status -like '*Ошибка в приказе*') {
            $label = '4 — Ошибка приказа: проверить исправление'
        } elseif ($changedOther -or $largeShift -or $operation -match 'компенсирована|Совпадение ФИО') {
            $label = '2 — Логическое закрытие: проверить основание'
        } else {
            $label = '1 — Наиболее безопасно: выравнивание до 1 дня'
        }
        if (-not $counts.ContainsKey($label)) { $counts[$label] = 0 }
        $counts[$label]++
        for ($offset = 0; $offset -lt 4; $offset++) {
            $sheet.Cells.Item(($first+$offset),19).Value2 = [string]$label
        }
    }
    $sheet.Range('S10:S1230').HorizontalAlignment = -4108
    $sheet.Range('S10:S1230').VerticalAlignment = -4108
    $sheet.Range('S10:S1230').WrapText = $true
    $excel.CalculateFull()
    $book.Save()
    $book.Close($false)
    $check = $excel.Workbooks.Open($outputPath,0,$true)
    Write-Output $check.Worksheets.Item(1).Range('A3').Text
    if ($check.Worksheets.Item(1).ListObjects.Item('FourRowReview').ListColumns.Count -ne 19) { throw 'Колонка не добавлена в таблицу' }
    $counts.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Output ($_.Name + ': ' + $_.Value + ' пар') }
    $check.Close($false)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}