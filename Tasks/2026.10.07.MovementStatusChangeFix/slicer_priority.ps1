$ErrorActionPreference = 'Stop'
$outputPath = Join-Path $PSScriptRoot 'Анализ смен статусов v24.xlsx'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Анализ смен статусов v23.xlsx') -Destination $outputPath
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open($outputPath)
    $types = $book.SlicerCaches.Item(1)
    $statuses = $book.SlicerCaches.Item(2)
    # Типы всегда доступны; статусы без данных скрываются.
    $types.CrossFilterType = 1
    $statuses.CrossFilterType = 4
    $types.ClearManualFilter()
    $statuses.ClearManualFilter()
    $sheet = $book.Worksheets.Item(1)
    if ($sheet.FilterMode) { $sheet.ShowAllData() }
    $typeName = [string]$types.SlicerItems.Item(1).Name
    foreach ($item in $types.SlicerItems) { $item.Selected = ($item.Name -eq $typeName) }
    $excel.CalculateFull()
    $available = @($statuses.SlicerItems | Where-Object { $_.HasData })
    Write-Output ('Доступных статусов для первого типа: ' + $available.Count + ' из ' + $statuses.SlicerItems.Count)
    if ($available.Count -eq 0) { throw 'Нет доступных статусов' }
    $statusName = [string]$available[0].Name
    foreach ($item in $statuses.SlicerItems) { $item.Selected = ($item.Name -eq $statusName) }
    $excel.CalculateFull()
    Write-Output ('Совместный фильтр: ' + $sheet.Range('A3').Text)
    $types.ClearManualFilter()
    $statuses.ClearManualFilter()
    $excel.CalculateFull()
    $book.Save()
    $book.Close($false)
    $check = $excel.Workbooks.Open($outputPath, 0, $true)
    Write-Output ('Режимы после открытия: ' + $check.SlicerCaches.Item(1).CrossFilterType + ', ' + $check.SlicerCaches.Item(2).CrossFilterType)
    Write-Output $check.Worksheets.Item(1).Range('A3').Text
    $check.Close($false)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}