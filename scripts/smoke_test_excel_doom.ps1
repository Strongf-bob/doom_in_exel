param(
    [string]$WorkbookPath = ""
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $root "output\spreadsheet\ExcelDoom.xlsm"
}

if (-not (Test-Path $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($WorkbookPath)

    $excel.Run("'$($workbook.Name)'!ExcelDoom_StartGame")
    $excel.Run("'$($workbook.Name)'!ExcelDoom_MoveForward")
    $excel.Run("'$($workbook.Name)'!ExcelDoom_TurnRight")
    $excel.Run("'$($workbook.Name)'!ExcelDoom_Shoot")
    $excel.Run("'$($workbook.Name)'!ExcelDoom_StopGame")
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "Smoke test passed for $WorkbookPath"
