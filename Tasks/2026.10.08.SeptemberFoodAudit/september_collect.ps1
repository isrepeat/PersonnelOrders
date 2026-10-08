$ErrorActionPreference = 'Stop'
$septemberDir = 'C:/WORK/Windows/Строевые приказы/Tasks/2026.10.08.SeptemberFoodAudit/Results'
[void][IO.Directory]::CreateDirectory($septemberDir)
foreach ($septemberName in @('РУХ_last.xlsx','Продовольче.xlsm')) {
    $septemberSource = [IO.File]::Open("C:/Users/isrepeat/OneDrive/Рабочий стол/$septemberName", [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    try {
        $septemberTarget = [IO.File]::Create("$septemberDir/$septemberName")
        try { $septemberSource.CopyTo($septemberTarget) } finally { $septemberTarget.Dispose() }
    } finally { $septemberSource.Dispose() }
}
$septemberExcel = New-Object -ComObject Excel.Application
$septemberExcel.Visible = $false
$septemberExcel.DisplayAlerts = $false
$septemberExcel.EnableEvents = $false
$septemberExcel.AutomationSecurity = 3
$septemberBook = $null
try {
    $septemberBook = $septemberExcel.Workbooks.Open("$septemberDir/РУХ_last.xlsx", 0, $false)
    $septemberExcel.Calculation = -4135
    $septemberSheet = $septemberBook.Worksheets.Item('Статистика')
    $septemberExcel.CalculateFullRebuild()
    $septemberResults = @()
    for ($septemberDay = 1; $septemberDay -le 30; $septemberDay++) {
        $septemberDate = [DateTime]::new(2026, 9, $septemberDay)
        $septemberInput = [object[,]]::new(1, 1)
        $septemberInput.SetValue($septemberDate.ToOADate(), 0, 0)
        $septemberSheet.Range('B3').Value2 = $septemberInput
        $septemberExcel.Calculate()
        $septemberValues = $septemberSheet.Range('C4:H8').Value2
        $septemberResults += [pscustomobject]@{
            date = $septemberDate.ToString('yyyy-MM-dd')
            food = $septemberValues.GetValue(5, 2)
            cooked = $septemberValues.GetValue(5, 3)
            dry = $septemberValues.GetValue(5, 4)
            removed = $septemberValues.GetValue(5, 5)
            enrolled = $septemberValues.GetValue(5, 6)
            other = $septemberValues.GetValue(4, 2)
        }
        Write-Output ("Calculated " + $septemberDate.ToString('dd.MM') + ': ' + ($septemberResults[-1] | ConvertTo-Json -Compress))
    }
    $septemberJson = $septemberResults | ConvertTo-Json -Depth 5 -Compress
    $septemberOut = [IO.File]::Create("$septemberDir/statistics.json.gz")
    $septemberGzip = [IO.Compression.GZipStream]::new($septemberOut, [IO.Compression.CompressionMode]::Compress)
    $septemberWriter = [IO.StreamWriter]::new($septemberGzip, [Text.UTF8Encoding]::new($false))
    try { $septemberWriter.Write($septemberJson) } finally { $septemberWriter.Dispose(); $septemberGzip.Dispose(); $septemberOut.Dispose() }
    $septemberBook.Close($false)
    $septemberBook = $null
} finally {
    if ($null -ne $septemberBook) { $septemberBook.Close($false) }
    $septemberExcel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($septemberExcel)
}