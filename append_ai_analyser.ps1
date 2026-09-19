param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$DataPath,
    [Parameter(Mandatory = $true)][string]$SheetName
)

$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
$openedHere = $false

try {
    $rows = Get-Content -Raw -LiteralPath $DataPath | ConvertFrom-Json
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
}

try {
    foreach ($candidate in $excel.Workbooks) {
        if ([string]::Equals($candidate.FullName, (Resolve-Path $WorkbookPath).Path, [StringComparison]::OrdinalIgnoreCase)) {
            $workbook = $candidate
            break
        }
    }

    if ($null -eq $workbook) {
        $workbook = $excel.Workbooks.Open((Resolve-Path $WorkbookPath).Path, 0, $false)
        $openedHere = $true
    }

    $excel.DisplayAlerts = $false
    $sheet = $workbook.Worksheets.Add()
    $sheet.Name = $SheetName
    $missing = [System.Type]::Missing
    $sheet.Move($missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
    $sheet.Tab.Color = 49407

    for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
        $row = $rows[$rowIndex]
        for ($columnIndex = 0; $columnIndex -lt $row.Count; $columnIndex++) {
            $value = $row[$columnIndex]
            if ($null -eq $value) { $value = '' }
            $cell = $sheet.Cells.Item($rowIndex + 1, $columnIndex + 1)
            if ($value.PSObject.Properties.Name -contains 'formula') {
                $cell.Formula = [string]$value.formula
            } elseif ($value.PSObject.Properties.Name -contains 'value') {
                $cell.Value2 = [double]$value.value
            } else {
                $cell.Value2 = [string]$value
            }
        }
    }

    $sheet.Rows.Item(1).Font.Bold = $true
    $sheet.Rows.Item(4).Font.Bold = $true
    $sheet.Rows.Item(5).Font.Bold = $true

    $yellow = 65535
    $descriptionColumn = 2
    $dateColumn = 4
    if ($SheetName -like 'AI Journal*') {
        $sheet.Cells.Item(1, 3).NumberFormat = 'dd/mm/yyyy'

        for ($rowIndex = 4; $rowIndex -le $rows.Count; $rowIndex++) {
            $description = [string]$sheet.Cells.Item($rowIndex, $descriptionColumn).Value2
            if ($description -eq 'GC Fees' -or $description -eq 'Income Allocation') {
                $sheet.Rows.Item($rowIndex).Interior.Color = $yellow
            }

            if ($sheet.Cells.Item($rowIndex, $dateColumn).Value2 -ne $null -and [string]$sheet.Cells.Item($rowIndex, $dateColumn).Value2 -ne '') {
                $sheet.Cells.Item($rowIndex, $dateColumn).NumberFormat = 'dd/mm/yyyy'
                $sheet.Cells.Item($rowIndex, $dateColumn).Interior.Color = $yellow
            }
        }
    }

    $sheet.Columns.Item('A:J').AutoFit()
    $workbook.Save()
    Write-Output "Created '$SheetName' without rewriting existing worksheets."
} finally {
    if ($workbook -and $openedHere) { $workbook.Close($true) }
    if ($excel -and $openedHere) { $excel.Quit() }
}