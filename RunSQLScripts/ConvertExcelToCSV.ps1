
Function ExcelToCsv {
    [CmdletBinding()] 
    param(
        [ Parameter(Mandatory = $true) ] [string] $FolderName,
        [ Parameter(Mandatory = $true) ] [string] $InputFile,
        [ Parameter(Mandatory = $true) ] [string] $OutputFileWoExt, # 'FileNameWoExt'
        [ Parameter(Mandatory = $true) ] [string] $PostFix # '_Sheet'
    )
    $excelFile = Join-Path $FolderName $InputFile
    $ExcelObj = New-Object -ComObject Excel.Application
    $wb = $ExcelObj.Workbooks.Open($excelFile)
    
    $count = 0
    foreach ($ws in $wb.Worksheets) {
        $count++
        $sheetName = $PostFix + $count 
        $csvFileFullPath = $FolderName + "\" + $OutputFileWoExt + $sheetName + ".csv"
        if ((test-path $csvFileFullPath)) {
            Write-Host "Replace previous csv file: "$csvFileFullPath -ForegroundColor Magenta
            Remove-Item $csvFileFullPath -Force
        }
        Write-Output $csvFileFullPath 
        $ws.SaveAs($csvFileFullPath, 6)

    }
    $ExcelObj.Workbooks.Close()
    $ExcelObj.Quit()
}


<#
$myFolder = "C:\Users\gazho\OneDrive - Microsoft\Zhou_Data\Projects_Assets_IP_and_Tools\PowerShell_Scripts\Run_Scripts_With_Stress_Testing"
$myExcelFile = "Stress-Testing-Config.xlsx"
$myCsvFileWoExt = "Stress-Testing-Config"
$mySheetName = "-Sheet"


$myExcelFileFullPath = Join-Path $myFolder $myExcelFile 

$myCsvFileWoExt = [System.IO.Path]::GetFileNameWithoutExtension($myExcelFileFullPath)


ExcelToCsv -FolderName $myFolder -InputFile $myExcelFile -OutputFileWoExt $myCsvFileWoExt -PostFix $mySheetName

#>