param([Parameter(Mandatory = $true)][string]$OutputPath)
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class FoodHistoryExcel {
    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object result);
    public static object ReadExcel() {
        var clsid = new Guid("00024500-0000-0000-C000-000000000046");
        object result;
        GetActiveObject(ref clsid, IntPtr.Zero, out result);
        return result;
    }
}
'@
$excelHistory = [FoodHistoryExcel]::ReadExcel()
$targetHistory = [IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $targetHistory) { throw 'Копия истории уже существует; выберите новое имя' }
$desktopHistory = Join-Path ([Environment]::GetFolderPath('Desktop')) 'РУХ_last.xlsx'
foreach ($bookHistory in $excelHistory.Workbooks) {
    if ($bookHistory.FullName -eq $desktopHistory) {
        $bookHistory.SaveCopyAs($targetHistory)
        Write-Output "Копия открытого РУХ сохранена: $targetHistory"
        return
    }
}
throw 'Настольный РУХ_last не найден в открытом Excel'