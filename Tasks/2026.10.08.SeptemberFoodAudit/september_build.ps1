$ErrorActionPreference = 'Stop'
$reportDir = 'C:/WORK/Windows/Строевые приказы/Tasks/2026.10.08.SeptemberFoodAudit/Results'
$reportPath = "$reportDir/Продовольче — сверка сентября 2026.xlsm"
$reportInput = [IO.File]::OpenRead("$reportDir/september.json.gz")
$reportGzip = [IO.Compression.GZipStream]::new($reportInput, [IO.Compression.CompressionMode]::Decompress)
$reportReader = [IO.StreamReader]::new($reportGzip, [Text.Encoding]::UTF8)
try { $reportData = $reportReader.ReadToEnd() | ConvertFrom-Json } finally { $reportReader.Dispose(); $reportGzip.Dispose(); $reportInput.Dispose() }
Copy-Item -LiteralPath "$reportDir/Продовольче.xlsm" -Destination $reportPath -Force
function Report-Color([int]$r, [int]$g, [int]$b) { return $r + 256 * $g + 65536 * $b }
function Report-Row($values, [int]$width) {
    $matrix = [object[,]]::new(1, $width)
    for ($column = 0; $column -lt $width; $column++) {
        $item = $values[$column]
        if ($item -is [long] -or $item -is [int]) { $item = [double]$item }
        $matrix.SetValue($item, 0, $column)
    }
    return ,$matrix
}
$reportExcel = New-Object -ComObject Excel.Application
$reportExcel.Visible = $false
$reportExcel.DisplayAlerts = $false
$reportExcel.EnableEvents = $false
$reportExcel.AutomationSecurity = 3
$reportAutofill = $reportExcel.AutoCorrect.AutoFillFormulasInLists
$reportExcel.AutoCorrect.AutoFillFormulasInLists = $false
$reportBook = $null
try {
    $reportBook = $reportExcel.Workbooks.Open($reportPath, 0, $false)
    $reportExcel.Calculation = -4135
    $reportSheet = $reportBook.Worksheets.Item('Котел')
    $reportTable = $reportSheet.ListObjects.Item('ПЗ')
    if ($reportSheet.FilterMode) { $reportSheet.ShowAllData() }
    # Снимки текста сохраняются без запуска пользовательской VBA-функции MillByQty.
    $reportText = [object[,]]::new($reportData.sourceResults.Count, 1)
    for ($reportIndex = 0; $reportIndex -lt $reportData.sourceResults.Count; $reportIndex++) { $reportText.SetValue($reportData.sourceResults[$reportIndex], $reportIndex, 0) }
    $reportSheet.Range("AS3:AS$($reportData.sourceResults.Count + 2)").Value2 = $reportText
    foreach ($reportDay in ($reportData.days | Sort-Object sourceRow -Descending)) {
        $reportRow = [int]$reportDay.sourceRow
        if ($reportDay.changes.Count -gt 0) {
            [void]$reportTable.ListRows.Add($reportRow - 1)
            $reportNewRow = $reportRow + 1
            $reportSheet.Range("A${reportRow}:AU${reportRow}").Copy($reportSheet.Range("A${reportNewRow}:AU${reportNewRow}"))
            $reportSheet.Range("A${reportRow}:AQ${reportRow}").Value2 = (Report-Row $reportDay.old 43)
            $reportSheet.Range("A${reportNewRow}:AQ${reportNewRow}").Value2 = (Report-Row $reportDay.new 43)
            $reportPair = $reportSheet.Range("A${reportRow}:AU${reportNewRow}")
            $reportPair.FormatConditions.Delete()
            $reportPair.Font.Color = (Report-Color 255 255 255)
            $reportPair.Font.Name = 'Arial'
            $reportPair.Font.Size = 10
            $reportPair.RowHeight = 24
            $reportSheet.Range("A${reportRow}:AU${reportRow}").Interior.Color = (Report-Color 46 43 39)
            $reportSheet.Range("A${reportNewRow}:AU${reportNewRow}").Interior.Color = (Report-Color 84 72 57)
            foreach ($reportColumn in $reportDay.changes) {
                $reportSheet.Cells.Item($reportRow, [int]$reportColumn).Interior.Color = (Report-Color 255 102 102)
                $reportSheet.Cells.Item($reportNewRow, [int]$reportColumn).Interior.Color = (Report-Color 146 208 80)
                $reportSheet.Cells.Item($reportRow, [int]$reportColumn).Font.Bold = $true
                $reportSheet.Cells.Item($reportNewRow, [int]$reportColumn).Font.Bold = $true
            }
            $reportSheet.Cells.Item($reportRow, 47).Value2 = "ИСХОДНОЕ — строка $reportRow исходного файла. " + [string]$reportDay.notes
            $reportSheet.Cells.Item($reportNewRow, 47).Value2 = 'ПЕРЕСЧИТАНО — ' + [string]$reportDay.notes
            $reportSheet.Cells.Item($reportNewRow, 45).Value2 = 'Строка сверки с исправленными числовыми данными. Текст приказа в этой сравнительной копии не формировался.'
            if ($reportDay.ambiguous) {
                foreach ($reportMarkedRow in @($reportRow, $reportNewRow)) {
                    foreach ($reportColumn in @(26,42)) { $reportSheet.Cells.Item($reportMarkedRow, $reportColumn).Interior.Color = (Report-Color 255 217 102) }
                }
            }
        }
    }
    $reportSheet.UsedRange.Font.Name = 'Arial'
    $reportSheet.UsedRange.Font.Size = 10
    $reportSheet.Range('A1:AU2').Interior.Color = (Report-Color 84 72 57)
    $reportSheet.Range('A1:AU2').Font.Color = (Report-Color 255 255 255)
    $reportSheet.Range('A1:AU2').Font.Bold = $true
    $reportSheet.Range('A2:AU2').RowHeight = 42
    $reportSheet.Columns.Item(47).ColumnWidth = 65
    $reportControl = $reportBook.Worksheets.Add()
    $reportControl.Name = 'Контроль сентября'
    $reportControl.Move([Type]::Missing, $reportBook.Worksheets.Item($reportBook.Worksheets.Count))
    $reportControl.Range('A1').Value2 = 'Сентябрь 2026 — сверка продовольственного обеспечения'
    $reportControl.Range('A1').Font.Size = 16
    $reportControl.Range('A1').Font.Bold = $true
    $reportNotes = @(
        'Красная строка — исходные значения; зелёная под ней — пересчитанные. Яркие ячейки — значения к замене. Жёлтый — спорный источник.',
        'Все 30 дней пересчитаны в Microsoft Excel по формулам Статистика!B3, D8:H8 и D7 на копии актуального РУХ_last. Исходные файлы не изменены.',
        'Сухе — из РУХ; СП. ВСЬОГО и подразделения — независимый подсчёт персональных назначений. Начало включительно, прекращение не включается.',
        '№257 исправлен: 50 человек назначены на 05–07 сентября; подсчёт выполнен по обновлённому тексту приказа.',
        'Столбцы ГІ+Полігон…УПР на продовольчому сохранены: их историческая разбивка не проверялась. ІП пересчитано как остаток только при наличии этой разбивки.',
        '01–10 сентября отсутствовали исходные показатели РУХ, а Сухе суммировало итог вместе с подразделениями. Данные восполнены; неизвестное ІП оставлено пустым.',
        '27 сентября: по №279 исправлена разбивка 11 человек СБ вместо 2ТБ, 2 человека МР вместо ІП. Численность по приказам — по ФИО, а не печатным итогам.',
        'Проверка отражает актуальные записи РУХ, а не архивные снимки на каждый день. Строки пары нельзя суммировать; Результат исходных строк — снимок, для новых текст приказа не формировался.'
    )
    for ($reportIndex = 0; $reportIndex -lt $reportNotes.Count; $reportIndex++) {
        $reportControl.Cells.Item($reportIndex + 2, 1).Value2 = $reportNotes[$reportIndex]
    }
    $reportHeaders = @('Дата','Продовольче','Котлове','Сухе РУХ','Знято','Зараховано','ЗІЧ','СП по приказам (рабочее)','СП буквально по тексту','РУХ − рабочее СП','Изменённых ячеек','Приказы','Примечание')
    $reportControl.Range('A11:M11').Value2 = (Report-Row $reportHeaders 13)
    $reportControlValues = [object[,]]::new(30, 13)
    for ($reportIndex = 0; $reportIndex -lt 30; $reportIndex++) {
        $reportDay = $reportData.days[$reportIndex]
        $reportStats = $reportDay.stats
        $reportValues = @($reportDay.new[0], $reportStats.food, $reportStats.cooked, $reportStats.dry, $reportStats.removed, $reportStats.enrolled, $reportStats.other, $reportDay.workingOrders, $reportDay.strictOrders, ($reportStats.dry - $reportDay.workingOrders), $reportDay.changes.Count, ($reportDay.orderNumbers -join ', '), $reportDay.notes)
        for ($reportColumn = 0; $reportColumn -lt 13; $reportColumn++) {
            $reportItem = $reportValues[$reportColumn]
            if ($reportItem -is [long] -or $reportItem -is [int]) { $reportItem = [double]$reportItem }
            $reportControlValues.SetValue($reportItem, $reportIndex, $reportColumn)
        }
    }
    $reportControl.Range('A12:M41').Value2 = $reportControlValues
    [void]$reportControl.ListObjects.Add(1, $reportControl.Range('A11:M41'), $null, 1)
    $reportControl.Range('A12:A41').NumberFormat = $reportSheet.Range('A410').NumberFormat
    $reportControl.Range('A11:M41').WrapText = $true
    $reportControl.Range('A11:M41').Font.Name = 'Arial'
    $reportControl.Range('A11:M41').Font.Size = 10
    $reportControl.Range('A11:L41').ColumnWidth = 16
    $reportControl.Columns.Item(13).ColumnWidth = 85
    $reportControl.Range('A11:M11').RowHeight = 42
    $reportControl.Range('A12:M41').RowHeight = 72
    $reportControl.Range('A1:M41').Interior.Color = (Report-Color 46 43 39)
    $reportControl.Range('A1:M41').Font.Color = (Report-Color 255 255 255)
    $reportControl.Range('A11:M11').Interior.Color = (Report-Color 84 72 57)
    $reportControl.Range('A11:M11').Font.Bold = $true
    for ($reportStyleRow = 12; $reportStyleRow -le 41; $reportStyleRow += 2) {
        $reportControl.Range("A${reportStyleRow}:M${reportStyleRow}").Interior.Color = (Report-Color 84 72 57)
    }
    $reportControl.Range('A1:M9').RowHeight = 22
    for ($reportIndex = 0; $reportIndex -lt 30; $reportIndex++) {
        if ($reportData.days[$reportIndex].ambiguous) { $reportControl.Range("H$($reportIndex+12):I$($reportIndex+12)").Interior.Color = (Report-Color 255 217 102) }
    }
    $reportSheet.Activate()
    $reportStart = [DateTime]::new(2026,9,1).ToOADate()
    $reportStop = [DateTime]::new(2026,10,1).ToOADate()
    $reportTable.Range.AutoFilter(1, ">=$reportStart", 1, "<$reportStop")
    $reportBook.Windows.Item(1).FreezePanes = $false
    $reportBook.Windows.Item(1).Zoom = 70
    $reportExcel.Goto($reportSheet.Range('A410'), $true)
    $reportExcel.CutCopyMode = 0
    $reportExcel.Calculation = -4105
    $reportExcel.Calculate()
    $reportFinalRow = [int]$reportData.days[-1].sourceRow + [int]$reportData.summary.changedDays
    $reportSheet.Range("A410:AU${reportFinalRow}").RowHeight = 24
    $reportBook.Save()
    Write-Output ("Saved: " + $reportPath + '; days=' + $reportData.summary.changedDays + '; cells=' + $reportData.summary.changedCells)
    $reportBook.Close($false)
    $reportBook = $null
} finally {
    if ($null -ne $reportBook) { $reportBook.Close($false) }
    $reportExcel.AutoCorrect.AutoFillFormulasInLists = $reportAutofill
    $reportExcel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reportExcel)
}