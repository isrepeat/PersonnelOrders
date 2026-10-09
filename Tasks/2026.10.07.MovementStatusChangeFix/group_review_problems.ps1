$ErrorActionPreference = 'Stop'
$outputPath = Join-Path $PSScriptRoot 'Анализ смен статусов v31.xlsx'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Анализ смен статусов v30.xlsx') -Destination $outputPath
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open($outputPath)
    $calculation = $excel.Calculation
    $excel.Calculation = -4135
    $sheet = $book.Worksheets.Item(1)
    foreach ($cache in $book.SlicerCaches) { $cache.ClearManualFilter() }
    if ($sheet.FilterMode) { $sheet.ShowAllData() }
    $destination = 11
    $groups = @{}
    $keptIds = @{}
    for ($first = 11; $first -le 354; $first += 4) {
        if ($sheet.Cells.Item($first,19).Text -notlike '5 —*') { continue }
        $operation = [string]$sheet.Cells.Item($first,14).Value2
        $status = [string]$sheet.Cells.Item($first,16).Value2
        # Основная группа отражает причину, мешающую подтвердить закрытие.
        if ($status -like '*отсутствует приказ*') { $group = 'Приказ отсутствует' }
        elseif ($operation -like '*Соответствующее событие не найдено*') { $group = 'Не найдено соответствующее событие в приказе' }
        elseif ($operation -like '*действие и прежний статус пропущены*') { $group = 'В пункте пропущены действие и прежний статус' }
        elseif ($operation -like '*амбулаторный или стационарный характер*') { $group = 'Неоднозначный тип лечения: стационар или ВЛК' }
        elseif ($operation -like '*прежний статус назван*') { $group = 'Противоречие прежнего статуса в РУХе и приказе' }
        elseif ($status -like '*однозначная дата начала отпуска*') { $group = 'Не установлена однозначная дата начала отпуска' }
        elseif ($status -like '*дата выписки*' -and $status -like '*дата справки ВЛК*') { $group = 'Не установлены даты выписки и ВЛК' }
        elseif ($status -like '*дата выписки*') { $group = 'Не установлена дата выписки' }
        elseif ($status -like '*дата справки ВЛК*') { $group = 'Не установлена дата справки ВЛК' }
        else { $group = 'Иное противоречие: проверить пункт вручную' }
        if ($first -ne $destination) { [void]$sheet.Range('A'+$first+':U'+($first+3)).Copy($sheet.Range('A'+$destination+':U'+($destination+3))) }
        $keptIds[[string]$sheet.Cells.Item($destination,20).Value2] = $true
        $groups[$destination] = $group
        $destination += 4
    }
    $last = $destination - 1
    $sheet.Range('A'+$destination+':Z354').Clear()
    $sheet.Range('A'+$destination+':Z1600').Interior.Color = 0x252729
    [void]$sheet.Columns.Item('T').Insert()
    $sheet.Columns.Item('T').Hidden = $false
    $sheet.Columns.Item('U:V').Hidden = $true
    $sheet.ListObjects.Item('FourRowReview').Resize($sheet.Range('A10:T'+$last))
    [void]$sheet.Range('S10:S'+$last).Copy($sheet.Range('T10:T'+$last))
    $sheet.Range('T10').Value2 = 'Группа проблемы'
    $sheet.Columns.Item('T').ColumnWidth = 48
    $counts = @{}
    foreach ($row in $groups.Keys) {
        $group = [string]$groups[$row]
        if (-not $counts.ContainsKey($group)) { $counts[$group] = 0 }
        $counts[$group]++
        for ($offset = 0; $offset -lt 4; $offset++) { $sheet.Cells.Item(([int]$row+$offset),20).Value2 = $group }
    }
    $sheet.Range('T10:T'+$last).WrapText = $true
    $sheet.Range('T10:T'+$last).HorizontalAlignment = -4108
    $sheet.Range('T10:T'+$last).VerticalAlignment = -4108
    $sheet.Range('A3').Formula2 = ([string]$sheet.Range('A3').Formula2).Replace('354',[string]$last)
    $sheet.Range('A5').Value2 = 'Остались случаи уровня 5. «Группа проблемы» — основная причина; дополнительные сомнения сохранены в «Операция». Исправленные пары уровня 4 исключены.'
    # Сохраняем на втором листе только оставшиеся пары, без уже закрытого уровня 4.
    $audit = $book.Worksheets.Item(2)
    if ($audit.FilterMode) { $audit.ShowAllData() }
    $destination = 6
    for ($row = 6; $row -le 91; $row++) {
        if (-not $keptIds.ContainsKey([string]$audit.Cells.Item($row,1).Value2)) { continue }
        if ($row -ne $destination) { [void]$audit.Range('A'+$row+':M'+$row).Copy($audit.Range('A'+$destination+':M'+$destination)) }
        $destination++
    }
    $audit.Range('A'+$destination+':Z400').Clear()
    $audit.Range('A'+$destination+':Z400').Interior.Color = 0x252729
    $audit.ListObjects.Item('LeaveStartRuleCheck').Resize($audit.Range('A5:M'+($destination-1)))
    $audit.Range('A2').Value2 = '82 оставшихся перехода требуют ручной проверки. Основные причины сгруппированы на первом листе в колонке «Группа проблемы».'
    $excel.Calculation = $calculation
    $excel.CalculateFull()
    $book.Save()
    $book.Close($false)
    $check = $excel.Workbooks.Open($outputPath,0,$true)
    Write-Output $check.Worksheets.Item(1).Range('A3').Text
    Write-Output ('Столбцов таблицы: '+$check.Worksheets.Item(1).ListObjects.Item('FourRowReview').ListColumns.Count)
    $counts.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Output ($_.Name+': '+$_.Value+' пар') }
    $check.Close($false)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}