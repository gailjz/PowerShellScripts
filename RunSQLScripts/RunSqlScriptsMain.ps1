############################################################################################################
############################################################################################################
#
# Author: Gaiye "Gail" Zhou
# May 2020
# Description: Scripts to run SQL Scripts stored in files 
#   with an option of running stress testing against the SQL Server DB (Or Azure Synapse DB)
#   one user runs selected queries multiple times (configurable)
############################################################################################################
Function GetPassword([SecureString] $securePassword) {
    $securePassword = Read-Host "Enter Password" -AsSecureString
    $P = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))
    return $P
}

##################################################################
#
# Main program starts here
#
##################################################################

$ScriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
Set-Location -Path $ScriptPath

. "$ScriptPath\ConvertExcelToCSV.ps1"  #use this if using Powershell ver 2
# "$PSCommandPath\ConvertExcelToCSV.ps1" #use this if using Powershell ver 3+

. "$ScriptPath\ExecuteScriptFile.ps1"  #use this if using Powershell ver 2
# "$PSCommandPath\ExecuteScriptFile.ps1" #use this if using Powershell ver 3+

$error.Clear()

$displayMsg = " Program starts now. Be prepared to confirm or specify the configuration file. "
Write-Host " "
Write-Host $displayMsg -ForegroundColor Cyan 
Write-Host " "

$defaultCfgFilePath = $ScriptPath 
$cfgFilePath = Read-Host -prompt "Enter the Config File Path or press 'Enter' to accept the default [$($defaultCfgFilePath)]"
if ([string]::IsNullOrEmpty($cfgFilePath)) {
    $cfgFilePath = $ScriptPath
}

$defaultCfgFile = "ScriptConfig-1.xlsx"
#$defaultCfgFile = "ScriptConfig-2.xlsx"
$cfgFile = Read-Host -prompt "Enter the Config File Name or press 'Enter' to accept the default [$($defaultCfgFile)]"
if ([string]::IsNullOrEmpty($cfgFile)) {
    $cfgFile = $defaultCfgFile
}
$CfgFileFullPath = join-path $cfgFilePath $cfgFile
if (!(test-path $CfgFileFullPath )) {
    Write-Host "Could not find Config File: $CfgFileFullPath " -ForegroundColor Red
    break 
}


#=====================================================
# Get SQL Server Security Information 
#=====================================================
# Get User Name and Password 
$defaultIntegrated = "Yes"
$Integrated = Read-Host -prompt "Enter 'Yes' or 'No' to connect using integrated Security or press 'Enter' to accept the default [$($defaultIntegrated)]"
if ([string]::IsNullOrEmpty($Integrated)) {
    $Integrated = $defaultIntegrated
}

if ($Integrated.toUpper() -eq "NO") {
    Write-Host "Please Enter SQLAUTH Login Information..." -ForegroundColor Yellow
    $UserName = Read-Host -prompt "Enter the User Name"
    if ([string]::IsNullOrEmpty($UserName)) {
        Write-Host "A user name must be entered" -ForegroundColor Red
        break
    }
    $Password = GetPassword
    if ([string]::IsNullOrEmpty($Password)) {
        Write-Host "A password must be entered." -ForegroundColor Red
        break
    }

}


#=============================================================================================
# Convert xslx file into two csv files and use them to extract information 
# CSV file generated used below naming convention
#    Xslx_FileName_Without_Ext-Sheet1.csv
#    Xslx_FileName_Without_Ext-Sheet2.csv
#=============================================================================================

$ExcelObj = New-Object -ComObject Excel.Application;
$Workbook = $ExcelObj.WorkBooks.Open($CfgFileFullPath);

$CsvFileWoExt = [System.IO.Path]::GetFileNameWithoutExtension($CfgFileFullPath)
$SheetName = "-Sheet"
# Call Function to Convert Excel File Sheets into CSV files 
ExcelToCsv -FolderName $cfgFilePath -InputFile $cfgFile -OutputFileWoExt $CsvFileWoExt -PostFix $SheetName
# Make fure the configuration file sheets match below arrangements: 
$serverDbCfgFile = $CsvFileWoExt + $SheetName + '1' + ".csv"

$serverDbCfgCsv = Import-Csv $serverDbCfgFile

$server = "localMachine\GailzSqlSvr2017"
# $server = "'.\GAILZSQLSVR2017"
$database = "AdventureWorksDW2017"
$connectionTimeOut = 30
$queryTimeOut = 30
$workersCount = 1 # just set an initial value. To be overwritten from Config File 

ForEach ($csvItem in $serverDbCfgCsv) {
    $server = $csvItem.ServerName
    $database = $csvItem.DatabaseName
    $connectionTimeOut = $csvItem.ConnectionTimeOut
    $queryTimeOut = $csvItem.QueryTimeOut
    $workersCount = $csvItem.WorkersCount
    $msg = "Server Name: " + $server + " Database Name: " + $database + "WorkersCount: " + $workersCount
    Write-Host $msg -ForegroundColor Blue
}

$error.Clear()

# Store Sheet 2 and greater into individual csv files
$number = 0; 
$arrayHash = @()
foreach ($sheet in $Workbook.Worksheets) {
    $number++
    if ($number -ge 2) {
        $scriptCfgConfigFileName = $CsvFileWoExt + $SheetName + $number + ".csv"
        $scriptsCfgFileFullPath = join-path $cfgFilePath $scriptCfgConfigFileName
        $statusLogFileName = $scriptCfgConfigFileName = $CsvFileWoExt + $SheetName + $number + "-Log" + ".csv"
        $statusLogFileFullPath = join-path $cfgFilePath $statusLogFileName
        $workBookHashTableItem = @{ }
        $workBookHashTableItem.Clear()
        $workBookHashTableItem.add("ScriptCfgFileFullPath", $scriptsCfgFileFullPath)
        $workBookHashTableItem.add("StatusLogFileFullPath", $statusLogFileFullPath)
        $arrayHash += $workBookHashTableItem
    }

}
$ExcelObj.Workbooks.Close()
$ExcelObj.Quit()


$error.Clear()

#=====================================================
# Get SQL Server Connection 
#=====================================================


#$MySqlConnection = New-Object System.Data.SqlClient.SqlConnection("Data Source=$server;Integrated Security=SSPI;Initial Catalog=$database;)
$MySqlConnection = New-Object System.Data.SqlClient.SqlConnection("Data Source=$server;Integrated Security=SSPI;Initial Catalog=$database;Connect Timeout=$connectionTimeOut")

if ($Integrated.toUpper() -eq 'NO') {
    $MySqlConnection = New-Object System.Data.SqlClient.SqlConnection("Data Source=$server;Integrated Security=false;Initial Catalog=$database;User ID=$UserName;Password=$Password;Connect Timeout=$connectionTimeOut")    
}
$MySqlConnection.open()

#======================================================================
# Run Scripts Based on Each Sheet of Configuration - 
#=======================================================================
if ($workersCount -eq 1) {

    Foreach ($arrayItem in $arrayHash) {
        $scriptsCfgFileFullPath = $arrayItem.ScriptCfgFileFullPath
        $statusLogFileFullPath = $arrayItem.StatusLogFileFullPath

        if ((test-path $statusLogFileFullPath)) {
            Write-Host "Replace previous status log file: "$statusLogFileFullPath -ForegroundColor Magenta
            Remove-Item $statusLogFileFullPath -Force
        }

        $HeaderRow = "Active", "ScriptType", "ScriptFileFolder", "ScriptFileName", "Variables", "NumberExec", "PauseTimeInSec", "Status", "DurationSec"
        $HeaderRow -join ","  >> $statusLogFileFullPath

        ProcessConfigAndRunScript -csvFileFullPath $scriptsCfgFileFullPath -statusLogFileFullPath $statusLogFileFullPath -Connection $MySqlConnection -queryTimeout $queryTimeOut
    }
}
# This is not working yet. It is designed to run the scripts tasks in parallel 
elseif ($workersCount -ge 2) {
    # Run Scripts Based on Each Sheet of Configuration - Multi-Threaded
    Write-Host "This part of the program is not done yet. Make sure your WorkersCount is equal to 1. Your WorkersCount is "$workersCount"." -ForegroundColor Red
    workflow ExeParallel {

        foreach -parallel ($arrayItem in $arrayHash) {
      
            $scriptsCfgFileFullPath = $arrayItem.ScriptCfgFileFullPath
            $statusLogFileFullPath = $arrayItem.StatusLogFileFullPath
    
            if ((test-path $statusLogFileFullPath)) {
                #Write-Host "Replace previous status log file: "$statusLogFileFullPath -ForegroundColor Magenta
                Remove-Item $statusLogFileFullPath -Force
            }
    
            $HeaderRow = "Active", "ScriptType", "ScriptFileFolder", "ScriptFileName", "Variables", "NumberExec", "PauseTimeInSec", "Status", "DurationSec"
            #$HeaderRow -join ","  >> $statusLogFileFullPath
    
            ProcessConfigAndRunScript -csvFileFullPath $scriptsCfgFileFullPath -statusLogFileFullPath $statusLogFileFullPath -Connection $MySqlConnection -queryTimeout $queryTimeOut
        }
    
    }
    ExeParallel
    
}
else {
    Write-Host "Received unexpected value in your config file. WorkerCount: " $workersCount -ForegroundColor red 
}

#====================================================
# Close SQL Connection 
$MySqlConnection.Close()
#====================================================

$FinishTime = Get-Date 
Write-Host "Finished work at " $FinishTime  
Write-Host "Have a great day!" 
