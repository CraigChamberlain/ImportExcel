#Requires -Modules Pester

if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

Describe "Close-Excel via a filestream" {
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

    It "Should not throw an exception" {
        $stream = [System.IO.File]::Open($xlfileImportColumns, [System.IO.FileMode]::Open)
        Open-ExcelPackage -stream $stream | Close-ExcelPackage 

        Open-ExcelPackage -stream $stream | Close-ExcelPackage -Show

        Open-ExcelPackage -stream $stream | Close-ExcelPackage -NoSave

        $ExcelPath = ([System.IO.Path]::GetTempFileName() -replace "\.tmp", ".xlsx") 
        Open-ExcelPackage -stream $stream | Close-ExcelPackage -SaveAs $ExcelPath
        $ExcelPath | Should -Exist
        Remove-Item $ExcelPath
        $stream.Dispose()
    }

    It "Behaves Properly" {
        $bytes = [System.IO.File]::ReadAllBytes("$xlfileImportColumns")
        $stream = [System.IO.MemoryStream]::new($bytes)
        $stream2 = [System.IO.MemoryStream]::new()
        $pkg = Open-ExcelPackage -stream $stream 
        $pkg.File | Should -BeNullOrEmpty
        
        #todo add a page 
        $pkg = @{Name = 1} | Export-Excel -ExcelPackage $pkg -WorksheetName "NewSheet" -PassThru
        $pkg.SaveAs($stream2)
        $stream.Dispose()
        #open edited stream
  
        $pkg = Open-ExcelPackage -stream $stream2 
        

        #todo check for additional sheet
        $pkg.Workbook.Worksheets.Name | Should -Contain "NewSheet"
        $pkg | Close-ExcelPackage -NoSave
        $stream2.Dispose()


    }

    It "Behaves Badly" {
        $bytes = [System.IO.File]::ReadAllBytes("$xlfileImportColumns")
        $stream = [System.IO.MemoryStream]::new($bytes)
        $pkg = Open-ExcelPackage -stream $stream 
        $pkg.File | Should -BeNullOrEmpty
        
        #todo add a page 
        $pkg = @{Name = 1} | Export-Excel -ExcelPackage $pkg -WorksheetName "NewSheet" -PassThru
        $pkg | Close-ExcelPackage
        
        #open edited stream
  
        $pkg = Open-ExcelPackage -stream $stream
        

        #todo check for additional sheet
        $pkg.Workbook.Worksheets.Name | Should -Contain "NewSheet"
        $pkg | Close-ExcelPackage -NoSave
        $stream.Dispose()


    }

    It "Behaves Badly 2" {
        $stream = [System.IO.File]::Open($xlfileImportColumns, [System.IO.FileMode]::Open)
        $pkg = Open-ExcelPackage -stream $stream 
        $pkg.File | Should -BeNullOrEmpty
        
        #todo add a page 
        $pkg = @{Name = 1} | Export-Excel -ExcelPackage $pkg -WorksheetName "NewSheet" -PassThru
        $pkg | Close-ExcelPackage
        $stream.Dispose()
        #open edited file
  
        $pkg = Open-ExcelPackage -Path $xlfileImportColumns
        

        #todo check for additional sheet
        $pkg.Workbook.Worksheets.Name | Should -Contain "NewSheet"
        $pkg | Close-ExcelPackage -NoSave

    }
    It "Behaves Well 2" {
        $stream = [System.IO.File]::Open($xlfileImportColumns, [System.IO.FileMode]::Open)
        $pkg = Open-ExcelPackage -stream $stream 
        $pkg.File | Should -BeNullOrEmpty
        
        #todo add a page 
        $pkg = @{Name = 1} | Export-Excel -ExcelPackage $pkg -WorksheetName "NewSheet" -PassThru
        $pkg | Close-ExcelPackage
        $stream.Dispose()
        #open edited file
  
        $pkg = Open-ExcelPackage -Path $xlfileImportColumns
        

        #todo check for additional sheet
        $pkg.Workbook.Worksheets.Name | Should -Contain "NewSheet"
        $pkg | Close-ExcelPackage -NoSave

    }

}