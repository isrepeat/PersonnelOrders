Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class FoodExcelReference {
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
$excelFoodReference = [FoodExcelReference]::ReadExcel()
$targetFoodReference = Join-Path (Resolve-Path 'Tasks/01.Food/Work/2026.10.09').Path 'РУХ_last_format_reference.xlsx'
$foundFoodReference = $false
foreach ($bookFoodReference in $excelFoodReference.Workbooks) {
    if ($bookFoodReference.Name -eq 'РУХ_last.xlsx') {
        $bookFoodReference.SaveCopyAs($targetFoodReference)
        $foundFoodReference = $true
        Write-Output 'Reference snapshot saved; original workbook unchanged'
        break
    }
}
if (-not $foundFoodReference) { throw 'РУХ_last not found in running Excel instance' }