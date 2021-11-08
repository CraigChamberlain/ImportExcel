#Requires -Modules Pester
$scriptPath = $PSScriptRoot
Import-Module $scriptPath\..\..\ImportExcel.psd1 -Force
$tfp = "$scriptPath\Read-OleDbData.xlsx"
$IsMissingACE = $null -eq ((New-Object system.data.oledb.oledbenumerator).GetElements().SOURCES_NAME -like "Microsoft.ACE.OLEDB*")
if($IsMissingACE){
    Write-Host "MICROSOFT.ACE.OLEDB is missing! Tests will be skipped. Please see https://www.microsoft.com/en-us/download/details.aspx?id=54920"
}
Describe "Invoke-ExcelQuery" -Tag "Invoke-ExcelQuery" {
    $PSDefaultParameterValues = @{ 'It:Skip' = $IsMissingACE }
    Context "Sheet1`$A1" {
        It "Should return 1 result with a value of 1" {
            $Results = Invoke-ExcelQuery $tfp "select ROUND(F1) as [A1] from [sheet1`$A1:A1]"
            @($Results).length + $Results.A1 | Should -Be 2
        }
    }
}
