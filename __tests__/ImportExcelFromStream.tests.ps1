#Requires -Modules Pester

if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

Describe "Import-Excel via a filestream" {
    BeforeAll {

        $xlfileImportColumns = "$PSScriptRoot\testImportExcelImportColumns.xlsx"

        # Create $xlfileImportColumns if it does not exist
        if (!(Test-Path -Path $xlfileImportColumns)) {
            $xl = "" | Export-Excel $xlfileImportColumns -PassThru

            Set-ExcelRange -Worksheet $xl.Sheet1 -Range A1 -Value 'A'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range B1 -Value 'B'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range C1 -Value 'C'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range D1 -Value 'D'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range E1 -Value 'E'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range F1 -Value 'F'
    
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range A2 -Value '1'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range B2 -Value '2'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range C2 -Value '3'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range D2 -Value '4'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range E2 -Value '5'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range F2 -Value '6'
    
            Close-ExcelPackage $xl
        }
    }

    AfterAll {
        Remove-Item $PSScriptRoot\testImportExcelSparse.xlsx -ErrorAction SilentlyContinue
    }

    It "Should import as a stream" {
        $stream = [System.IO.File]::Open($xlfileImportColumns, [System.IO.FileMode]::Open)
        $actual = @(Import-Excel -stream $stream)
        $actualNames = $actual[0].psobject.properties.Name

        $actualNames.Count | Should -Be 6
        $actualNames[0] | Should -Be 'A'
        $actualNames[2] | Should -Be 'C'

        $actual.Count | Should -Be 1
        $actual[0].A | Should -Be 1
        $actual[0].B | Should -Be 2
        $actual[0].C | Should -Be 3
        $actual[0].D | Should -Be 4
        $actual[0].E | Should -Be 5
    }

}