$ErrorActionPreference = 'Stop'
$transferDir = 'C:/WORK/Windows/Строевые приказы/outputs/dry-rations-audit-20261008'
$targetPath = 'C:/Users/isrepeat/OneDrive/Рабочий стол/РУХ_last.xlsx'
$transferInput = [IO.File]::OpenRead("$transferDir/transfer.json.gz")
$transferGzip = [IO.Compression.GZipStream]::new($transferInput, [IO.Compression.CompressionMode]::Decompress)
$transferReader = [IO.StreamReader]::new($transferGzip, [Text.Encoding]::UTF8)
try { $transferData = $transferReader.ReadToEnd() | ConvertFrom-Json } finally { $transferReader.Dispose(); $transferGzip.Dispose(); $transferInput.Dispose() }
$transferExcel = New-Object -ComObject Excel.Application
$transferExcel.Visible = $false
$transferExcel.DisplayAlerts = $false
$transferExcel.EnableEvents = $false
$transferBook = $null
$transferAutoFill = $transferExcel.AutoCorrect.AutoFillFormulasInLists
$transferExcel.AutoCorrect.AutoFillFormulasInLists = $false
try {
    $transferBook = $transferExcel.Workbooks.Open($targetPath, 0, $false)
    if ($transferBook.ReadOnly) { throw 'РУХ_last открыт только для чтения; изменения не выполнены.' }
    $transferSheet = $transferBook.Worksheets.Item('Сухпрод')
    $transferTable = $transferSheet.ListObjects.Item('СУХПРОД')
    if ($transferTable.Range.Address($false, $false) -ne $transferData.tableRef) { throw 'Диапазон СУХПРОД изменился после подготовки.' }
    $transferExcel.Calculation = -4135
    $transferMetadataLast = $transferTable.Range.Row + $transferTable.Range.Rows.Count - 1
    $transferMetadata = $transferSheet.Range("H3:I${transferMetadataLast}").Value2
    foreach ($transferUpdate in $transferData.updates) {
        $transferRow = [int]$transferUpdate.row
        if ($transferSheet.Cells.Item($transferRow, 3).Text -ne $transferUpdate.name) { throw "ФИО изменилось в строке $transferRow" }
        $transferMetadata.SetValue([string]$transferUpdate.basis, $transferRow - 2, 1)
        $transferMetadata.SetValue([double]$transferUpdate.order, $transferRow - 2, 2)
        $transferSheet.Cells.Item($transferRow, 9).NumberFormat = '0'
        if ($transferUpdate.note) {
            $transferSheet.Cells.Item($transferRow, 10).Value2 = [string]$transferUpdate.note
            $transferSheet.Range("B${transferRow}:J${transferRow}").Interior.Color = 7876720
            $transferSheet.Cells.Item($transferRow, 10).WrapText = $true
        }
    }
    $transferSheet.Range("H3:I${transferMetadataLast}").Value2 = $transferMetadata
    $transferLastRow = $transferTable.Range.Row + $transferTable.Range.Rows.Count - 1
    $transferNewLast = $transferLastRow + $transferData.additions.Count
    if ($transferData.additions.Count -gt 0) {
        for ($transferCheckRow = $transferLastRow + 1; $transferCheckRow -le $transferNewLast; $transferCheckRow++) {
            if ($transferExcel.WorksheetFunction.CountA($transferSheet.Range("B${transferCheckRow}:J${transferCheckRow}")) -ne 0) { throw "Данные ниже таблицы в строке $transferCheckRow" }
        }
        $transferTable.Resize($transferSheet.Range("B2:J${transferNewLast}"))
        foreach ($transferValues in $transferData.additions) {
            $transferLastRow++
            $transferSheet.Range("B$($transferLastRow-1):J$($transferLastRow-1)").Copy()
            $transferSheet.Range("B${transferLastRow}:J${transferLastRow}").PasteSpecial(-4122)
            $transferRowValues = [object[,]]::new(1, 9)
            for ($transferCol = 0; $transferCol -lt 9; $transferCol++) {
                $transferValue = $transferValues[$transferCol]
                if ($transferCol -eq 3 -or $transferCol -eq 5) { $transferValue = [DateTime]::ParseExact($transferValue, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture).ToOADate() }
                if ($transferValue -is [long] -or $transferValue -is [int]) { $transferValue = [double]$transferValue }
                $transferRowValues.SetValue($transferValue, 0, $transferCol)
            }
            $transferSheet.Range("B${transferLastRow}:J${transferLastRow}").Value2 = $transferRowValues
            # Существующие форматы дат копируются из книги без новых formatCode.
            $transferSheet.Cells.Item($transferLastRow, 5).NumberFormat = $transferSheet.Cells.Item(3, 5).NumberFormat
            $transferSheet.Cells.Item($transferLastRow, 7).NumberFormat = $transferSheet.Cells.Item(3, 7).NumberFormat
            $transferSheet.Cells.Item($transferLastRow, 7).Formula = "=E${transferLastRow}+F${transferLastRow}"
            $transferSheet.Cells.Item($transferLastRow, 9).NumberFormat = '0'
            $transferSheet.Range("B${transferLastRow}:J${transferLastRow}").Interior.Color = 4218161
            $transferSheet.Range("B${transferLastRow}:J${transferLastRow}").Font.Color = 16777215
        }
    }
    $transferExcel.CutCopyMode = 0
    $transferExcel.Calculation = -4105
    $transferBook.Save()
    Write-Output ("Saved: updates=" + $transferData.updates.Count + "; additions=" + $transferData.additions.Count + "; table=" + $transferTable.Range.Address($false, $false))
    $transferBook.Close($false)
    $transferBook = $null
} finally {
    if ($null -ne $transferBook) { $transferBook.Close($false) }
    $transferExcel.AutoCorrect.AutoFillFormulasInLists = $transferAutoFill
    $transferExcel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($transferExcel)
}