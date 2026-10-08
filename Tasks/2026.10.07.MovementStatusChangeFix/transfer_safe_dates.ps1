param([string]$Levels = '1')
$ErrorActionPreference = 'Stop'
$env:PYTHONIOENCODING = 'utf-8'
$python = 'C:/Users/isrepeat/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
$plan = (& $python -W ignore (Join-Path $PSScriptRoot 'prepare_safe_transfer.py') $Levels) | ConvertFrom-Json
if ($LASTEXITCODE -ne 0 -or $plan.conflicts.Count) { throw 'Есть несопоставленные исходные строки; перенос остановлен' }
$sourcePath = 'C:/Users/isrepeat/OneDrive/Рабочий стол/РУХ_last.xlsx'
$baselinePath = 'C:/Users/isrepeat/OneDrive/Рабочий стол/РУХ_2026.10.08.xlsx'
$backupPath = Join-Path $PSScriptRoot ('РУХ_last до исправлений уровней ' + $Levels.Replace(',', '-') + ' ' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.xlsx')
Copy-Item -LiteralPath $sourcePath -Destination $backupPath
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open($sourcePath, 0, $false)
    if ($book.ReadOnly) { $book.Close($false); throw 'РУХ_last открыт только для чтения: требуется разблокировать файл' }
    $previousCalculation = $excel.Calculation
    $excel.Calculation = -4135
    $sheet = $book.Worksheets.Item('Відсутні')
    # Проверка всех исходных значений до первой записи предотвращает частичный перенос.
    foreach ($update in $plan.updates) {
        $before = [DateTime]::ParseExact($update.before, 'yyyy-MM-dd', $null)
        if ($sheet.Cells.Item([int]$update.row,[int]$update.column).Value2 -ne $before.ToOADate()) { throw 'Исходная дата изменилась после подготовки переноса' }
    }
    foreach ($update in $plan.updates) {
        $after = [DateTime]::ParseExact($update.after, 'yyyy-MM-dd', $null)
        $sheet.Cells.Item([int]$update.row,[int]$update.column).Formula = '=DATE(' + $after.Year + ',' + $after.Month + ',' + $after.Day + ')'
        # Сохраняем обычное значение даты, как в исходной таблице.
        $sheet.Cells.Item([int]$update.row,[int]$update.column).Value2 = [double]$after.ToOADate()
    }
    $excel.Calculation = $previousCalculation
    $excel.CalculateFull()
    $book.Save()
    $book.Close($false)
    $check = $excel.Workbooks.Open($sourcePath, 0, $true)
    foreach ($update in $plan.updates) {
        $after = [DateTime]::ParseExact($update.after, 'yyyy-MM-dd', $null)
        if ($check.Worksheets.Item('Відсутні').Cells.Item([int]$update.row,[int]$update.column).Value2 -ne $after.ToOADate()) { throw 'Сохранённая дата не совпала с исправлением' }
    }
    $check.Close($false)
    Write-Output ('Исправлено дат: ' + $plan.updates.Count)
    Write-Output ('Резервная копия: ' + $backupPath)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}
$report = Join-Path (Get-Location) ('Results/ExcelComparing/РУХ_last.status_levels_' + $Levels.Replace(',', '_') + '_vs_РУХ_2026.10.08.diff.html')
& $python 'Tools/compare_excel.py' $baselinePath $sourcePath -o $report
if ($LASTEXITCODE -ne 0) { throw 'Не удалось создать отчёт сравнения' }
Write-Output ('Отчёт: ' + $report)