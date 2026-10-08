$ErrorActionPreference = 'Stop'
$env:PYTHONIOENCODING = 'utf-8'
$analysis = (& 'C:/Users/isrepeat/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe' -W ignore (Join-Path $PSScriptRoot 'compensated_leave.py')) | ConvertFrom-Json
if ($LASTEXITCODE -ne 0) { throw 'Ошибка анализа сроков' }
$outputPath = Join-Path $PSScriptRoot 'Анализ смен статусов v25.xlsx'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Анализ смен статусов v24.xlsx') -Destination $outputPath
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open($outputPath)
    $sheet = $book.Worksheets.Item(1)
    $audit = $book.Worksheets.Item(2)
    foreach ($cache in $book.SlicerCaches) { $cache.ClearManualFilter() }
    if ($sheet.FilterMode) { $sheet.ShowAllData() }
    foreach ($case in $analysis.compensated) {
        $audit.Cells.Item([int]$case.audit_row, 11).Value2 = 'Соответствует: задержка компенсирована'
        $audit.Cells.Item([int]$case.audit_row, 11).Font.Color = 0xE6ECF0
        $audit.Cells.Item([int]$case.audit_row, 12).Value2 = [string]$case.reason
        $date = [DateTime]::ParseExact($case.start, 'yyyy-MM-dd', $null)
        $first = [int]$case.block
        if ($first) {
            for ($offset = 0; $offset -lt 4; $offset++) {
                $row = $first + $offset
                $sheet.Cells.Item($row, 16).ClearContents()
                $field = if ($offset % 2 -eq 0) { 'Прибуття' } else { 'Вибуття' }
                $sheet.Cells.Item($row, 14).Value2 = [string]($field + ': ' + $date.ToString('dd.MM.yyyy') + ".`n`n" + $case.reason)
            }
            # Выравниваем только исправленную пару; исходные данные сохранены.
            foreach ($entry in @(@(($first + 2), 10), @(($first + 3), 8))) {
                $row = [int]$entry[0]
                $column = [int]$entry[1]
                $sheet.Cells.Item($row, $column).Formula = '=DATE(' + $date.Year + ',' + $date.Month + ',' + $date.Day + ')'
                $original = $sheet.Cells.Item(($row - 2), $column).Value2
                if ($original -ne $date.ToOADate()) {
                    $sheet.Cells.Item($row, $column).Interior.Color = 0x237D3B
                    $sheet.Cells.Item($row, $column).Font.Bold = $true
                }
            }
        }
        Write-Output ($case.name + ': ' + $case.days + ' + ' + $case.delta + ' = ' + $case.allowance)
    }
    $audit.Range('A2').Value2 = '290 переходов: 5 несоответствий, 197 совпадений, 2 задержки компенсированы сроком отпуска, 86 для проверки.'
    $audit.Range('A1').Value2 = 'Начало = max(выписка, ВЛК) + 1 день; задержка допустима, если срок + задержка = 30 / 45 / 60 / 120 дней.'
    $excel.CalculateFull()
    $book.Save()
    $book.Close($false)
    $check = $excel.Workbooks.Open($outputPath, 0, $true)
    foreach ($case in $analysis.compensated) {
        $first = [int]$case.block
        if ($check.Worksheets.Item(1).Cells.Item(($first + 2), 10).Value2 -ne $check.Worksheets.Item(1).Cells.Item(($first + 3), 8).Value2) { throw 'Даты перехода не совпадают' }
        if ($check.Worksheets.Item(1).Cells.Item($first, 16).Text) { throw 'Отметка ошибки не удалена' }
    }
    Write-Output ('Слайсеров: ' + $check.SlicerCaches.Count)
    Write-Output $check.Worksheets.Item(1).Range('A3').Text
    $check.Close($false)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}