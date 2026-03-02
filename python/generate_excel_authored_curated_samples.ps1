$ErrorActionPreference = "Stop"

$outDir = Resolve-Path "corpus/excel-authored/files"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Save-Xlsx {
    param(
        [Parameter(Mandatory = $true)] [object] $Workbook,
        [Parameter(Mandatory = $true)] [string] $Path
    )

    if (Test-Path $Path) {
        Remove-Item $Path -Force
    }
    # 51 = xlOpenXMLWorkbook (.xlsx)
    $Workbook.SaveAs($Path, 51)
    $Workbook.Close($true)
}

try {
    $formulaPath = Join-Path $outDir "internal-formula-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $ws.Range("A1").Value2 = 10
    $ws.Range("B1").Value2 = 20
    $ws.Range("C1").Formula = "=A1+B1"
    $ws.Range("D1").Value2 = "hello"
    $ws.Range("E1").Value2 = $true
    Save-Xlsx -Workbook $wb -Path $formulaPath

    $stylesPath = Join-Path $outDir "internal-styles-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $cell = $ws.Range("A1")
    $cell.Value2 = 1234.5
    $cell.NumberFormat = "#,##0.00"
    $cell.Font.Bold = $true
    $cell.Font.Color = 0x1F4E78
    $cell.Interior.Color = 0xD9F0E2
    $cell.HorizontalAlignment = -4108
    $ws.Range("A2").Value2 = "styled"
    Save-Xlsx -Workbook $wb -Path $stylesPath

    $commentsPath = Join-Path $outDir "internal-comments-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $ws.Range("A1").Value2 = "needs review"
    $ws.Range("A1").AddComment("Compatibility comment fixture")
    $ws.Range("B1").Value2 = 42
    Save-Xlsx -Workbook $wb -Path $commentsPath

    $chartPath = Join-Path $outDir "internal-chart-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $ws.Range("A1").Value2 = "Quarter"
    $ws.Range("B1").Value2 = "Revenue"
    $ws.Range("A2").Value2 = "Q1"
    $ws.Range("B2").Value2 = 12
    $ws.Range("A3").Value2 = "Q2"
    $ws.Range("B3").Value2 = 18
    $ws.Range("A4").Value2 = "Q3"
    $ws.Range("B4").Value2 = 9
    $ws.Range("A5").Value2 = "Q4"
    $ws.Range("B5").Value2 = 15
    $chartObject = $ws.ChartObjects().Add(220, 20, 420, 260)
    $chart = $chartObject.Chart
    $chart.SetSourceData($ws.Range("A1:B5"))
    $chart.ChartType = 51
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = "Quarterly Revenue"
    Save-Xlsx -Workbook $wb -Path $chartPath

    $definedNamesPath = Join-Path $outDir "internal-defined-names-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $ws.Range("A1").Value2 = "Amount"
    $ws.Range("A2").Value2 = 100
    $ws.Range("A3").Value2 = 200
    $ws.Range("A4").Value2 = 300
    $ws.Range("A5").Value2 = 400
    $ws.Range("C1").Value2 = "Total"
    $wb.Names.Add("SalesRange", "=Sheet1!`$A`$2:`$A`$5") | Out-Null
    $wb.Names.Add("TaxRate", "=0.07") | Out-Null
    $ws.Range("C2").Formula = "=SUM(SalesRange)"
    $ws.Range("C3").Formula = "=C2*TaxRate"
    Save-Xlsx -Workbook $wb -Path $definedNamesPath

    $tablePath = Join-Path $outDir "internal-table-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $ws.Range("A1").Value2 = "Region"
    $ws.Range("B1").Value2 = "Quarter"
    $ws.Range("C1").Value2 = "Sales"
    $ws.Range("A2").Value2 = "North"
    $ws.Range("B2").Value2 = "Q1"
    $ws.Range("C2").Value2 = 125
    $ws.Range("A3").Value2 = "North"
    $ws.Range("B3").Value2 = "Q2"
    $ws.Range("C3").Value2 = 142
    $ws.Range("A4").Value2 = "South"
    $ws.Range("B4").Value2 = "Q1"
    $ws.Range("C4").Value2 = 98
    $ws.Range("A5").Value2 = "South"
    $ws.Range("B5").Value2 = "Q2"
    $ws.Range("C5").Value2 = 111
    $table = $ws.ListObjects.Add(1, $ws.Range("A1:C5"), $null, 1)
    $table.Name = "SalesTable"
    $table.TableStyle = "TableStyleMedium2"
    Save-Xlsx -Workbook $wb -Path $tablePath

    $mergedCellsPath = Join-Path $outDir "internal-merged-cells-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $header = $ws.Range("A1:C1")
    $header.Merge()
    $header.Value2 = "Merged Header"
    $header.HorizontalAlignment = -4108
    $header.Font.Bold = $true
    $ws.Range("A2").Value2 = "Item"
    $ws.Range("B2").Value2 = "Qty"
    $ws.Range("C2").Value2 = "Price"
    $ws.Range("A3").Value2 = "Widget"
    $ws.Range("B3").Value2 = 5
    $ws.Range("C3").Value2 = 12.5
    Save-Xlsx -Workbook $wb -Path $mergedCellsPath

    $dataValidationPath = Join-Path $outDir "internal-data-validation-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $ws.Range("A1").Value2 = "Approval"
    $validationRange = $ws.Range("A2:A10")
    $validationRange.Validation.Delete()
    $validationRange.Validation.Add(3, 1, 1, "Yes,No")
    $validationRange.Validation.IgnoreBlank = $true
    $validationRange.Validation.InCellDropdown = $true
    $validationRange.Validation.InputTitle = "Pick value"
    $validationRange.Validation.InputMessage = "Choose Yes or No"
    $validationRange.Validation.ErrorTitle = "Invalid choice"
    $validationRange.Validation.ErrorMessage = "Value must be Yes or No"
    $validationRange.Validation.ShowError = $true
    $ws.Range("A2").Value2 = "Yes"
    Save-Xlsx -Workbook $wb -Path $dataValidationPath

    $conditionalFormattingPath = Join-Path $outDir "internal-conditional-formatting-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $ws.Range("A1").Value2 = "Score"
    $ws.Range("A2").Value2 = 72
    $ws.Range("A3").Value2 = 91
    $ws.Range("A4").Value2 = 64
    $ws.Range("A5").Value2 = 88
    $ruleRange = $ws.Range("A2:A5")
    $rule = $ruleRange.FormatConditions.Add(1, 5, "85")
    $rule.Font.Bold = $true
    $rule.Interior.Color = 0x99FF99
    Save-Xlsx -Workbook $wb -Path $conditionalFormattingPath

    $externalLinksPath = Join-Path $outDir "internal-external-links-baseline.xlsx"
    $externalLinkTargetPath = Join-Path $env:TEMP "rootcellar-external-link-target.xlsx"
    if (Test-Path $externalLinkTargetPath) {
        Remove-Item $externalLinkTargetPath -Force
    }

    try {
        $targetWb = $excel.Workbooks.Add()
        $targetWs = $targetWb.Worksheets.Item(1)
        $targetWs.Name = "Sheet1"
        $targetWs.Range("A1").Value2 = 321
        $targetWb.SaveAs($externalLinkTargetPath, 51)
        $targetWb.Close($true)

        $openedTargetWb = $excel.Workbooks.Open($externalLinkTargetPath)
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = "Sheet1"
        $ws.Range("A1").Value2 = "External Value"
        $ws.Range("A2").Formula = "='[$($openedTargetWb.Name)]Sheet1'!`$A`$1"
        Save-Xlsx -Workbook $wb -Path $externalLinksPath
        $openedTargetWb.Close($false)
    }
    finally {
        if (Test-Path $externalLinkTargetPath) {
            Remove-Item $externalLinkTargetPath -Force
        }
    }

    $pivotTablePath = Join-Path $outDir "internal-pivot-table-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $dataWs = $wb.Worksheets.Item(1)
    $dataWs.Name = "Data"
    $pivotWs = $wb.Worksheets.Add()
    $pivotWs.Name = "Pivot"
    $dataWs.Range("A1").Value2 = "Region"
    $dataWs.Range("B1").Value2 = "Quarter"
    $dataWs.Range("C1").Value2 = "Sales"
    $dataWs.Range("A2").Value2 = "North"
    $dataWs.Range("B2").Value2 = "Q1"
    $dataWs.Range("C2").Value2 = 120
    $dataWs.Range("A3").Value2 = "North"
    $dataWs.Range("B3").Value2 = "Q2"
    $dataWs.Range("C3").Value2 = 135
    $dataWs.Range("A4").Value2 = "South"
    $dataWs.Range("B4").Value2 = "Q1"
    $dataWs.Range("C4").Value2 = 95
    $dataWs.Range("A5").Value2 = "South"
    $dataWs.Range("B5").Value2 = "Q2"
    $dataWs.Range("C5").Value2 = 110
    $dataWs.Range("A6").Value2 = "West"
    $dataWs.Range("B6").Value2 = "Q1"
    $dataWs.Range("C6").Value2 = 102
    $dataWs.Range("A7").Value2 = "West"
    $dataWs.Range("B7").Value2 = "Q2"
    $dataWs.Range("C7").Value2 = 118

    $pivotCache = $wb.PivotCaches().Create(1, "'Data'!R1C1:R7C3")
    $pivotTable = $pivotCache.CreatePivotTable($pivotWs.Range("A3"), "SalesPivot")
    $rowField = $pivotTable.PivotFields("Region")
    $rowField.Orientation = 1
    $columnField = $pivotTable.PivotFields("Quarter")
    $columnField.Orientation = 2
    $pivotTable.AddDataField($pivotTable.PivotFields("Sales"), "Sum of Sales", -4157) | Out-Null
    Save-Xlsx -Workbook $wb -Path $pivotTablePath

    $queryConnectionPath = Join-Path $outDir "internal-query-connection-baseline.xlsx"
    $querySourcePath = Join-Path $env:TEMP "rootcellar-query-source.csv"
    "Name,Value`r`nAlpha,10`r`nBeta,20`r`nGamma,30" | Set-Content -Path $querySourcePath -Encoding UTF8
    try {
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = "Sheet1"
        $queryTable = $ws.QueryTables().Add("TEXT;$querySourcePath", $ws.Range("A1"))
        $queryTable.Name = "SourceQuery"
        # 1 = xlDelimited
        $queryTable.TextFileParseType = 1
        $queryTable.TextFileCommaDelimiter = $true
        # Disable background refresh so the connection/query parts are fully materialized before save.
        $queryTable.BackgroundQuery = $false
        $queryTable.Refresh($false) | Out-Null
        Save-Xlsx -Workbook $wb -Path $queryConnectionPath
    }
    finally {
        if (Test-Path $querySourcePath) {
            Remove-Item $querySourcePath -Force
        }
    }

    $sheetProtectionPath = Join-Path $outDir "internal-sheet-protection-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $ws.Range("A1").Value2 = "Locked"
    $ws.Range("A1").Locked = $true
    $ws.Range("B1").Value2 = "Editable"
    $ws.Range("B1").Locked = $false
    $ws.Protect("rootcellar")
    Save-Xlsx -Workbook $wb -Path $sheetProtectionPath

    $hyperlinksPath = Join-Path $outDir "internal-hyperlinks-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $ws.Range("A1").Value2 = "RootCellar"
    $ws.Hyperlinks().Add($ws.Range("A1"), "https://example.com/rootcellar") | Out-Null
    $ws.Range("A2").Value2 = "Jump to B5"
    $ws.Hyperlinks().Add($ws.Range("A2"), $null, "Sheet1!B5") | Out-Null
    $ws.Range("B5").Value2 = "Anchor Cell"
    Save-Xlsx -Workbook $wb -Path $hyperlinksPath

    $workbookProtectionPath = Join-Path $outDir "internal-workbook-protection-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $ws.Range("A1").Value2 = "Protected Structure"
    # Protect workbook structure (tabs cannot be added/moved/deleted without unprotect).
    $wb.Protect("rootcellar", $true, $false)
    Save-Xlsx -Workbook $wb -Path $workbookProtectionPath

    $printSettingsPath = Join-Path $outDir "internal-print-settings-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Sheet1"
    $ws.Range("A1").Value2 = "Item"
    $ws.Range("B1").Value2 = "Qty"
    $ws.Range("C1").Value2 = "Price"
    for ($i = 2; $i -le 30; $i++) {
        $ws.Range("A$i").Value2 = "Line $($i - 1)"
        $ws.Range("B$i").Value2 = $i - 1
        $ws.Range("C$i").Value2 = ($i - 1) * 2
    }
    # 2 = xlLandscape
    $ws.PageSetup.Orientation = 2
    $ws.PageSetup.Zoom = $false
    $ws.PageSetup.FitToPagesWide = 1
    $ws.PageSetup.FitToPagesTall = 1
    $ws.PageSetup.PrintTitleRows = "`$1:`$1"
    $ws.PageSetup.PrintArea = "`$A`$1:`$C`$30"
    Save-Xlsx -Workbook $wb -Path $printSettingsPath

    $calcChainPath = Join-Path $outDir "internal-calc-chain-baseline.xlsx"
    $wb = $excel.Workbooks.Add()
    $ws1 = $wb.Worksheets.Item(1)
    $ws1.Name = "Sheet1"
    $ws2 = $wb.Worksheets.Add()
    $ws2.Name = "Sheet2"
    $ws1.Range("A1").Value2 = 10
    $ws1.Range("A2").Value2 = 5
    $ws1.Range("B1").Formula = "=A1+A2"
    $ws1.Range("C1").Formula = "=Sheet2!A1+1"
    $ws2.Range("A1").Formula = "=Sheet1!B1*2"
    $ws2.Range("B1").Formula = "=SUM(A1,Sheet1!C1)"
    # Force calc chain regeneration for deterministic cross-sheet dependency metadata.
    $excel.CalculateFullRebuild()
    Save-Xlsx -Workbook $wb -Path $calcChainPath

    Write-Output "Generated curated Excel-authored baseline files in $outDir"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
