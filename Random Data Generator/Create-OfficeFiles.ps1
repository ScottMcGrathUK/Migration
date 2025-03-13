<#
.SYNOPSIS
    Generates test files (Word, Excel, and CSV) with large amounts of data.
    
.DESCRIPTION
    This script creates a specified number of Word, Excel, and CSV files with test data.
    To improve efficiency, it generates one template file per type and then duplicates it.

.PARAMETER NumWord
    Number of Word documents to create. Default is 10.

.PARAMETER NumExcel
    Number of Excel files to create. Default is 10.

.PARAMETER NumCSV
    Number of CSV files to create. Default is 10.

.EXAMPLE
    .\Generate-TestFiles.ps1 -NumWord 5 -NumExcel 5 -NumCSV 5

.NOTES
    Author: Scott McGrath
    Requires: Microsoft Office (for Word and Excel files)
    Run as administrator if permission issues arise.
#>

param(
    [int]$NumWord = 10,
    [int]$NumExcel = 10,
    [int]$NumCSV = 10
)

$basePath = "$env:USERPROFILE\Documents\TestFiles"
$wordPath = "$basePath\Word"
$excelPath = "$basePath\Excel"
$csvPath = "$basePath\CSV"

New-Item -Path $wordPath -ItemType Directory -Force | Out-Null
New-Item -Path $excelPath -ItemType Directory -Force | Out-Null
New-Item -Path $csvPath -ItemType Directory -Force | Out-Null

$wordTemplate = "$wordPath\DocumentTemplate.docx"
if (-not (Test-Path $wordTemplate)) {
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $doc = $wordApp.Documents.Add()
    $selection = $wordApp.Selection
    for ($i = 0; $i -lt 500; $i++) { 
        $selection.TypeText("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus luctus urna sed urna ultricies ac tempor dui sagittis.`r`n")
    }
    $doc.SaveAs($wordTemplate)
    $doc.Close()
    $wordApp.Quit()
}

$excelTemplate = "$excelPath\DocumentTemplate.xlsx"
if (-not (Test-Path $excelTemplate)) {
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $workbook = $excelApp.Workbooks.Add()
    $sheet = $workbook.Sheets.Item(1)
    for ($row = 1; $row -le 200; $row++) {
        for ($col = 1; $col -le 100; $col++) {
            $sheet.Cells.Item($row, $col) = "TestData"
        }
    }
    $workbook.SaveAs($excelTemplate)
    $workbook.Close($false)
    $excelApp.Quit()
}

$csvTemplate = "$csvPath\DocumentTemplate.csv"
if (-not (Test-Path $csvTemplate)) {
    $csvData = @()
    for ($row = 1; $row -le 200; $row++) {
        $csvData += ("TestData" * 100) -join ","
    }
    $csvData | Out-File -FilePath $csvTemplate -Encoding UTF8
}

for ($i = 1; $i -le $NumWord; $i++) {
    Copy-Item -Path $wordTemplate -Destination "$wordPath\Document$i.docx"
}

for ($i = 1; $i -le $NumExcel; $i++) {
    Copy-Item -Path $excelTemplate -Destination "$excelPath\Document$i.xlsx"
}

for ($i = 1; $i -le $NumCSV; $i++) {
    Copy-Item -Path $csvTemplate -Destination "$csvPath\Document$i.csv"
}

Write-Host "All files have been created successfully!"