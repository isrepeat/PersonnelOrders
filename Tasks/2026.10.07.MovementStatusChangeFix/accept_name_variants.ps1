$ErrorActionPreference = 'Stop'
$outputPath = Join-Path $PSScriptRoot 'Анализ смен статусов v27.xlsx'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Анализ смен статусов v26.xlsx') -Destination $outputPath
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open($outputPath)
    $sheet = $book.Worksheets.Item(1)
    foreach ($cache in $book.SlicerCaches) { $cache.ClearManualFilter() }
    if ($sheet.FilterMode) { $sheet.ShowAllData() }
    # Сравнение после приведения имени и отчества к именительному падежу.
    foreach ($case in @(
        @{ Row = 1127; Date = '2025-07-26'; Similarity = '96,4%'; Names = 'В РУХ: БОБРОВСЬКИЙ Євген Євгенович. В приказе: БОБРОВСЬКОГО Євгена Євгенійовича.' },
        @{ Row = 1211; Date = '2025-11-21'; Similarity = '90,9%'; Names = 'В РУХ: ЛОМАКО Сергій Іванович. В приказе: ЛОМАЦІ Сергію Івановичу.' }
    )) {
        $first = [int]$case.Row
        $date = [DateTime]::ParseExact($case.Date, 'yyyy-MM-dd', $null)
        for ($offset = 0; $offset -lt 4; $offset++) {
            $row = $first + $offset
            $sheet.Cells.Item($row, 2).Font.Color = 0x00FFFF
            $sheet.Cells.Item($row, 16).ClearContents()
            $text = [string]$sheet.Cells.Item($row, 13).Value2
            $text = $text.Replace("`nВозможное совпадение ПІБ: требуется подтверждение", '').Replace('Возможное совпадение ПІБ: требуется подтверждение', '')
            $sheet.Cells.Item($row, 13).Value2 = $text.TrimEnd()
            $operation = if ($offset % 2 -eq 0) { 'Прибуття' } else { 'Вибуття' }
            $sheet.Cells.Item($row, 14).Value2 = [string]($operation + ': ' + $date.ToString('dd.MM.yyyy') + ". Смена статуса: даты выровнены.`n`n" + $case.Names + ' Совпадение ФИО после учёта падежа: ' + $case.Similarity + ' (порог 85%); считаем одним человеком. ФИО выделено жёлтым из-за различий.' )
        }
        for ($offset = 2; $offset -lt 4; $offset++) {
            $row = $first + $offset
            $originalRow = $row - 2
            $sheet.Cells.Item($row, 4).Value2 = [string]$sheet.Cells.Item($originalRow, 4).Value2
            $sheet.Cells.Item($row, 4).Interior.Color = 0x3C5134
            $sheet.Cells.Item($row, 4).Font.Bold = $false
            $column = if ($offset -eq 2) { 10 } else { 8 }
            $sheet.Cells.Item($row, $column).Formula = '=DATE(' + $date.Year + ',' + $date.Month + ',' + $date.Day + ')'
            $changed = $sheet.Cells.Item($originalRow, $column).Value2 -ne $date.ToOADate()
            $sheet.Cells.Item($row, $column).Interior.Color = if ($changed) { 0x237D3B } else { 0x3C5134 }
            $sheet.Cells.Item($row, $column).Font.Bold = $changed
        }
    }
    $excel.CalculateFull()
    $book.Save()
    $book.Close($false)
    $check = $excel.Workbooks.Open($outputPath, 0, $true)
    foreach ($first in @(1127,1211)) {
        if ($check.Worksheets.Item(1).Cells.Item(($first + 2),10).Value2 -ne $check.Worksheets.Item(1).Cells.Item(($first + 3),8).Value2) { throw 'Даты не выровнены' }
        Write-Output ($check.Worksheets.Item(1).Cells.Item($first,2).Text + ': ' + $check.Worksheets.Item(1).Cells.Item(($first + 2),10).Text)
    }
    $check.Close($false)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}