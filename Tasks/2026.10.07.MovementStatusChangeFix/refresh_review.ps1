$ErrorActionPreference = 'Stop'
$env:PYTHONIOENCODING = 'utf-8'
$python = 'C:/Users/isrepeat/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
$plan = (& $python -W ignore (Join-Path $PSScriptRoot 'refresh_review.py')) | ConvertFrom-Json
if ($LASTEXITCODE -ne 0) { throw 'Не удалось проанализировать обновлённый РУХ' }
$outputPath = Join-Path $PSScriptRoot 'Анализ смен статусов v30.xlsx'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Анализ смен статусов v29.xlsx') -Destination $outputPath
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
function Set-ReviewValue($sheet,$row,$column,$value) {
    $cell = $sheet.Cells.Item([int]$row,[int]$column)
    if ($null -eq $value) { $cell.ClearContents() }
    elseif ($value -is [pscustomobject] -and $value.date) {
        $date = [DateTime]::ParseExact($value.date,'yyyy-MM-dd',$null)
        $cell.Value2 = [double]$date.ToOADate()
    } elseif ($value -is [long] -or $value -is [int] -or $value -is [double]) { $cell.Value2 = [double]$value }
    else { $cell.Value2 = [string]$value }
}
try {
    $book = $excel.Workbooks.Open($outputPath)
    $calculation = $excel.Calculation
    $excel.Calculation = -4135
    $sheet = $book.Worksheets.Item(1)
    $audit = $book.Worksheets.Item(2)
    foreach ($cache in $book.SlicerCaches) { $cache.ClearManualFilter() }
    if ($sheet.FilterMode) { $sheet.ShowAllData() }
    $destination = 11
    foreach ($case in $plan.selected) {
        $originalRow = [int]$case.old_row
        if ($originalRow -ne $destination) {
            [void]$sheet.Range('A'+$originalRow+':U'+($originalRow+3)).Copy($sheet.Range('A'+$destination+':U'+($destination+3)))
        }
        foreach ($edit in $case.edits) {
            $row = $destination + [int]$edit.offset
            Set-ReviewValue $sheet $row $edit.column $edit.value
            if ($edit.restore) {
                $sheet.Cells.Item($row,[int]$edit.column).Interior.Color = 0x3C5134
                $sheet.Cells.Item($row,[int]$edit.column).Font.Bold = $false
            }
        }
        $destination += 4
    }
    foreach ($case in $plan.new) {
        [void]$sheet.Range('A11:U14').Copy($sheet.Range('A'+$destination+':U'+($destination+3)))
        $reason = 'Приказ №'+$case.order+' от 06.10.2026 отсутствует в папке приказов. В РУХе: лечение закрыто 06.10.2026; отпуск начинается 07.10.2026. Проверить персональный пункт и число пунктов, дату выписки и справку ВЛК. Исправление пока не подтверждено; исходные статусы и даты сохранены.'
        for ($offset = 0; $offset -lt 4; $offset++) {
            $row = $destination + $offset
            $side = $offset % 2
            $sheet.Range('A'+$row+':P'+$row).ClearContents()
            $sheet.Range('A'+$row+':U'+$row).Interior.Color = if ($offset -lt 2) { 0x292D30 } else { 0x3C5134 }
            $sheet.Range('A'+$row+':U'+$row).Font.Color = 0xE6ECF0
            $sheet.Range('A'+$row+':U'+$row).Font.Bold = $false
            Set-ReviewValue $sheet $row 1 $case.type
            $columns = @(2,3,4,5,6,7,8,10,11,12)
            for ($index = 0; $index -lt $columns.Count; $index++) { Set-ReviewValue $sheet $row $columns[$index] $case.values[$side][$index] }
            Set-ReviewValue $sheet $row 9 $case.terms[$side]
            $sheet.Cells.Item($row,2).Font.Color = 0x66B3FF
            Set-ReviewValue $sheet $row 13 ('Пункт приказа не проверен: источник недоступен.'+"`n`n"+'Источник: 2026-10-06 №289.docx')
            Set-ReviewValue $sheet $row 14 $reason
            Set-ReviewValue $sheet $row 16 'Требует проверки | отсутствует приказ №289'
            $sheet.Cells.Item($row,17).Formula = '=SUBTOTAL(103,A'+$row+')'
            Set-ReviewValue $sheet $row 18 $(if ($offset -lt 2) {'Исходная'} else {'Исправленная'})
            Set-ReviewValue $sheet $row 19 '5 — Ручная проверка: данных недостаточно'
            Set-ReviewValue $sheet $row 20 $case.id
            Set-ReviewValue $sheet $row 21 $case.pair[$side]
            $sheet.Rows.Item($row).RowHeight = 180
        }
        $destination += 4
    }
    $last = $destination-1
    $sheet.Range('A'+$destination+':Z1230').Clear()
    $sheet.Range('A'+$destination+':Z1600').Interior.Color = 0x252729
    $sheet.ListObjects.Item('FourRowReview').Resize($sheet.Range('A10:S'+$last))
    $sheet.Range('A3').Formula2 = ([string]$sheet.Range('A3').Formula2).Replace('1230',[string]$last)
    $sheet.Range('A5').Value2 = 'Повторная проверка обновлённого РУХ_last: 09.10.2026. Исправленные переходы исключены; две новые пары требуют отсутствующего приказа №289.'
    $destination = 6
    $counts = @{}
    foreach ($originalRow in $plan.audit_rows) {
        if ($originalRow -ne $destination) { [void]$audit.Range('A'+$originalRow+':M'+$originalRow).Copy($audit.Range('A'+$destination+':M'+$destination)) }
        $result = [string]$audit.Cells.Item($destination,11).Text
        if (-not $counts.ContainsKey($result)) { $counts[$result]=0 }
        $counts[$result]++
        $destination++
    }
    foreach ($case in $plan.new) {
        [void]$audit.Range('A6:M6').Copy($audit.Range('A'+$destination+':M'+$destination))
        $audit.Range('A'+$destination+':M'+$destination).ClearContents()
        Set-ReviewValue $audit $destination 1 $case.id
        Set-ReviewValue $audit $destination 2 $case.name
        Set-ReviewValue $audit $destination 3 $case.values[0][1]
        Set-ReviewValue $audit $destination 4 $case.order
        Set-ReviewValue $audit $destination 11 'Требует проверки'
        $audit.Cells.Item($destination,11).Font.Color = 0xE6ECF0
        Set-ReviewValue $audit $destination 12 'Приказ №289 отсутствует. Проверить выписку, справку ВЛК, начало и срок отпуска; число персональных пунктов пока не установлено.'
        Set-ReviewValue $audit $destination 13 '2026-10-06 №289.docx'
        $destination++
    }
    $audit.Range('A'+$destination+':Z400').Clear()
    $audit.Range('A'+$destination+':Z400').Interior.Color = 0x252729
    $audit.ListObjects.Item('LeaveStartRuleCheck').Resize($audit.Range('A5:M'+($destination-1)))
    $audit.Range('A2').Value2 = [string](($destination-6).ToString()+' оставшихся переходов: '+$counts['Несоответствие правилу']+' несоответствия; остальные требуют проверки, включая 2 пары без приказа №289.')
    $excel.Calculation = $calculation
    $excel.CalculateFull()
    $book.Save()
    $book.Close($false)
    $check = $excel.Workbooks.Open($outputPath,0,$true)
    Write-Output $check.Worksheets.Item(1).Range('A3').Text
    Write-Output ('Строк данных: '+$check.Worksheets.Item(1).ListObjects.Item('FourRowReview').ListRows.Count)
    Write-Output ('Слайсеров: '+$check.SlicerCaches.Count)
    Write-Output $check.Worksheets.Item(2).Range('A2').Text
    $check.Close($false)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}