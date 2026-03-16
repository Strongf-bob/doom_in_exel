param(
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$modulePath = Join-Path $root "src\vba\modExcelDoom.bas"
$outputDir = Join-Path $root "output\spreadsheet"
$excelSecurityKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$bindingFlags = [System.Reflection.BindingFlags]

if (-not (Test-Path $modulePath)) {
    throw "VBA module not found: $modulePath"
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $outputDir "ExcelDoom.xlsm"
}

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$excel = $null
$workbook = $null
$securityValueExisted = $false
$previousAccessVBOM = $null

if (-not (Test-Path $excelSecurityKey)) {
    New-Item -Path $excelSecurityKey -Force | Out-Null
}

try {
    $existingValue = Get-ItemProperty -Path $excelSecurityKey -Name AccessVBOM -ErrorAction Stop
    $securityValueExisted = $true
    $previousAccessVBOM = $existingValue.AccessVBOM
}
catch {
    $securityValueExisted = $false
    $previousAccessVBOM = $null
}

try {
    Set-ItemProperty -Path $excelSecurityKey -Name AccessVBOM -Value 1 -Type DWord

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()

    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "DOOM"

    $vbProject = $workbook.GetType().InvokeMember("VBProject", $bindingFlags::GetProperty, $null, $workbook, $null)
    $vbComponents = $vbProject.GetType().InvokeMember("VBComponents", $bindingFlags::GetProperty, $null, $vbProject, $null)
    $moduleCode = Get-Content -Path $modulePath -Raw
    $moduleCode = [System.Text.RegularExpressions.Regex]::Replace($moduleCode, '^\s*Attribute VB_Name\s*=\s*".*?"\s*\r?\n', '', [System.Text.RegularExpressions.RegexOptions]::Multiline)
    $module = $vbComponents.GetType().InvokeMember("Add", $bindingFlags::InvokeMethod, $null, $vbComponents, @([int]1))
    $module.Name = "modExcelDoom"
    $module.CodeModule.AddFromString([string]$moduleCode)

    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }

    $xlOpenXMLWorkbookMacroEnabled = 52
    $workbook.SaveAs($OutputPath, $xlOpenXMLWorkbookMacroEnabled)
    $workbook.Close($false)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    $workbook = $null

    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($OutputPath)
    $excel.Run("ExcelDoom_ConfigureSheet")
    $workbook.Save()
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

    if ($securityValueExisted) {
        Set-ItemProperty -Path $excelSecurityKey -Name AccessVBOM -Value $previousAccessVBOM -Type DWord
    }
    else {
        Remove-ItemProperty -Path $excelSecurityKey -Name AccessVBOM -ErrorAction SilentlyContinue
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "Created $OutputPath"
