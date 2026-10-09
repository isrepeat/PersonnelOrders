param([string]$Mode = 'approved_groups')
$ErrorActionPreference = 'Stop'
$inputName = 'Анализ смен статусов v33.xlsx'
$outputName = 'Анализ смен статусов v34.xlsx'
$approved = @('Неоднозначный тип лечения: стационар или ВЛК', 'Не установлена однозначная дата начала отпуска')
if ($Mode -eq 'approved_medical_dates') {
    $inputName = 'Анализ смен статусов v34.xlsx'
    $outputName = 'Анализ смен статусов v35.xlsx'
    $approved = @('Не установлена дата выписки', 'Не установлены даты выписки и ВЛК', 'Не установлена дата справки ВЛК')
}
$outputPath = Join-Path $PSScriptRoot $outputName
Copy-Item -LiteralPath (Join-Path $PSScriptRoot $inputName) -Destination $outputPath
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
    $kept = @{}
    $destination = 11
    $previousLast = $sheet.ListObjects.Item('FourRowReview').ListRows.Count + 10
    $removed = 0
    for ($first = 11; $first -le $previousLast; $first += 4) {
        if ($sheet.Cells.Item($first,20).Text -in $approved) { $removed++; continue }
        if ($first -ne $destination) { [void]$sheet.Range('A'+$first+':V'+($first+3)).Copy($sheet.Range('A'+$destination+':V'+($destination+3))) }
        $kept[[string]$sheet.Cells.Item($destination,21).Value2] = $true
        $destination += 4
    }
    $last = $destination-1
    $sheet.Range('A'+$destination+':Z'+$previousLast).Clear()
    $sheet.Range('A'+$destination+':Z1600').Interior.Color = 0x252729
    $sheet.ListObjects.Item('FourRowReview').Resize($sheet.Range('A10:T'+$last))
    $sheet.Range('A3').Formula2 = ([string]$sheet.Range('A3').Formula2).Replace([string]$previousLast,[string]$last)
    $sheet.Range('A5').Value2 = [string]($removed.ToString()+' согласованных пар перенесены в РУХ_last и исключены. Остались случаи, требующие проверки.')
    $audit = $book.Worksheets.Item(2)
    if ($audit.FilterMode) { $audit.ShowAllData() }
    $destination = 6
    $auditLast = $audit.ListObjects.Item('LeaveStartRuleCheck').ListRows.Count + 5
    for ($row = 6; $row -le $auditLast; $row++) {
        if (-not $kept.ContainsKey([string]$audit.Cells.Item($row,1).Value2)) { continue }
        if ($row -ne $destination) { [void]$audit.Range('A'+$row+':M'+$row).Copy($audit.Range('A'+$destination+':M'+$destination)) }
        $destination++
    }
    $audit.Range('A'+$destination+':Z400').Clear()
    $audit.Range('A'+$destination+':Z400').Interior.Color = 0x252729
    $audit.ListObjects.Item('LeaveStartRuleCheck').Resize($audit.Range('A5:M'+($destination-1)))
    $audit.Range('A2').Value2 = [string]($kept.Count.ToString()+' оставшихся переходов требуют проверки. '+$removed+' согласованных пар перенесены в РУХ_last; статусы сохранены, даты связаны.')
    $excel.Calculation = $calculation
    $excel.CalculateFull()
    $book.Save()
    $book.Close($false)
    $check = $excel.Workbooks.Open($outputPath,0,$true)
    if ($check.Worksheets.Item(1).ListObjects.Item('FourRowReview').ListRows.Count -ne $kept.Count*4) { throw 'Неверное количество оставшихся строк' }
    for ($row = 11; $row -le $last; $row += 4) {
        if ($check.Worksheets.Item(1).Cells.Item($row,20).Text -in $approved) { throw 'Согласованная группа осталась в анализе' }
    }
    Write-Output $check.Worksheets.Item(1).Range('A3').Text
    Write-Output ('Слайсеров: '+$check.SlicerCaches.Count)
    $check.Close($false)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}