$ErrorActionPreference = 'Stop'
$taskDir = $PSScriptRoot
$sourcePath = Join-Path $taskDir 'Анализ смен статусов v22.xlsx'
$outputPath = Join-Path $taskDir 'Анализ смен статусов v23.xlsx'
Copy-Item -LiteralPath $sourcePath -Destination $outputPath
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open($outputPath)
    $sheet = $book.Worksheets.Item(1)
    $table = $sheet.ListObjects.Item('FourRowReview')
    # Два независимых слайсера применяют пересечение фильтров одной таблицы.
    $sheet.Range('6:9').RowHeight = 80
    $typeCache = $book.SlicerCaches.Item(1)
    $typeSlicer = $typeCache.Slicers.Item(1)
    $typeSlicer.Top = $sheet.Range('A6').Top + 5
    $typeSlicer.Left = 25
    $typeSlicer.Width = 1130
    $typeSlicer.Height = 140
    $typeSlicer.Style = 'StatusDarkRestored'
    $statusCache = $book.SlicerCaches.Add2($table, 'Статус пункта приказа', 'OrderPointStatusCache')
    $statusSlicer = $statusCache.Slicers.Add($sheet, [Type]::Missing, 'OrderPointStatusFilter', 'Статус пункта приказа', ($typeSlicer.Top + 145), 25, 1130, 165)
    $statusSlicer.NumberOfColumns = 2
    $statusSlicer.RowHeight = 32
    $statusSlicer.ColumnWidth = 555
    $statusSlicer.Style = 'StatusDarkRestored'
    $typeCache.ClearManualFilter()
    $statusCache.ClearManualFilter()
    if ($sheet.FilterMode) { $sheet.ShowAllData() }
    $excel.CalculateFull()
    # Проверяем совместное действие фильтров и реакцию счётчика.
    $typeItem = $typeCache.SlicerItems.Item(1)
    $typeName = [string]$typeItem.Name
    foreach ($item in $typeCache.SlicerItems) { $item.Selected = ($item.Name -eq $typeName) }
    $statusName = ''
    foreach ($item in $statusCache.SlicerItems) {
        if ($item.HasData -and $item.Name -ne '(blank)' -and $item.Name -ne '(пусто)' -and $item.Name -ne '') { $statusName = [string]$item.Name; break }
    }
    if ($statusName) {
        foreach ($item in $statusCache.SlicerItems) { $item.Selected = ($item.Name -eq $statusName) }
        $excel.CalculateFull()
        Write-Output ('Совместный фильтр: ' + $sheet.Range('A3').Text)
        for ($row = 11; $row -le 1230; $row++) {
            if (-not $sheet.Rows.Item($row).Hidden) {
                if ($sheet.Cells.Item($row, 1).Text -ne $typeName -or $sheet.Cells.Item($row, 16).Text -ne $statusName) { throw 'Совместная фильтрация не совпадает с выбранными значениями' }
            }
        }
    }
    $typeCache.ClearManualFilter()
    $statusCache.ClearManualFilter()
    $excel.CalculateFull()
    $book.Save()
    $book.Close($false)
    $check = $excel.Workbooks.Open($outputPath, 0, $true)
    Write-Output ('Слайсеров: ' + $check.SlicerCaches.Count)
    Write-Output ('Без фильтра: ' + $check.Worksheets.Item(1).Range('A3').Text)
    Write-Output ('Стиль: ' + $check.SlicerCaches.Item(2).Slicers.Item(1).Style.Name)
    $check.Close($false)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}