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

$defaultCfgFilePath = $ScriptPath

$displayMsg = " Program starts now. Be prepared to confirm or specify the configuration file. "
Write-Host " "
Write-Host $displayMsg -ForegroundColor Cyan 
Write-Host " "

$cfgFilePath = Read-Host -prompt "Enter the Config File Path or press 'Enter' to accept the default [$($defaultCfgFilePath)]"
if ([string]::IsNullOrEmpty($cfgFilePath)) {
    $cfgFilePath = $defaultCfgFilePath
}

$defaultCfgFile = "Run-SQLScripts-W-Stress-Config.xlsx"
$cfgFile = Read-Host -prompt "Enter the Config File Name or press 'Enter' to accept the default [$($defaultCfgFile)]"
if ([string]::IsNullOrEmpty($cfgFile)) {
    $cfgFile = $defaultCfgFile
}
$CfgFileFullPath = join-path $cfgFilePath $cfgFile
if (!(test-path $CfgFileFullPath )) {
    Write-Host "Could not find Config File: $CfgFileFullPath " -ForegroundColor Red
    break 
}
$StatusLogFileWoExt = [System.IO.Path]::GetFileNameWithoutExtension($CfgFileFullPath)
$StatusLogFile = $cfgFilePath + "\" + $StatusLogFileWoExt + "-Log.csv"
if ((test-path $StatusLogFile)) {
    Write-Host "Replace previous status log file: "$StatusLogFile -ForegroundColor Magenta
    Remove-Item $StatusLogFile -Force
}
Write-Host "Status Log File is: "$StatusLogFile -ForegroundColor Blue

$HeaderRow = "Active", "ScriptType", "ScriptFileFolder", "ScriptFileName", "Variables", "NumberExec", "PauseTimeInSec", "Status", "DurationSec"
$HeaderRow -join ","  >> $StatusLogFile


#=============================================================================================
# Convert xslx file into two csv files and use them to extract information 
# CSV file generated used below naming convention
#    Xslx_FileName_Without_Ext-Sheet1.csv
#    Xslx_FileName_Without_Ext-Sheet2.csv
#=============================================================================================

$CsvFileWoExt = [System.IO.Path]::GetFileNameWithoutExtension($CfgFileFullPath)
$SheetName = "-Sheet"
ExcelToCsv -FolderName $cfgFilePath -InputFile $cfgFile -OutputFileWoExt $CsvFileWoExt -PostFix $SheetName
$serverDbCfgFile = $CsvFileWoExt + $SheetName + '1' + ".csv"
$scriptsCfgFile = $CsvFileWoExt + $SheetName + '2' + ".csv"

$serverDbCfgCsv = Import-Csv $serverDbCfgFile
$scriptsCfgCsv = Import-Csv $scriptsCfgFile

#$server = "localMachine\GAILZSQLSVR2017"
$server = "'.\GAILZSQLSVR2017"
$database = "AdventureWorksDW2017"
$workersCount = 3 # this is default value 
ForEach ($csvItem in $serverDbCfgCsv) {
    $server = $csvItem.ServerName
    $database = $csvItem.DatabaseName
    $workersCount = $csvItem.WorkersCount
    $msg = "Server Name: " + $server + " Database Name: " + $database + "WorkersCount: " + $workersCount
    Write-Host $msg -ForegroundColor Blue
}


#=====================================================
# Get SQL Server Connection 
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

# default to use Integrated Security 
$MySqlConnection = New-Object System.Data.SqlClient.SqlConnection("Data Source=$server;Integrated Security=SSPI;Initial Catalog=$database;")

if ($Integrated.toUpper() -eq 'NO') {
    $MySqlConnection = New-Object System.Data.SqlClient.SqlConnection("Data Source=$server;Integrated Security=false;Initial Catalog=$database;User ID=$UserName;Password=$Password")

}


$error.Clear()

	
#Workflow RunParallelExecute {
#ForEach -Parallel -throttlelimit $workersCount ($S in $scriptsCfgCsv ) {
ForEach ($S in $scriptsCfgCsv ) {
    $StartDate = (Get-Date)
    $Active = $S.Active
    $rows = New-Object PSObject 
    if ($Active -eq '1') {
        $ScriptType = $S.ScriptType
        $ScriptFileFolder = $S.ScriptFileFolder
        $ScriptFileName = $S.ScriptFileName
        $Variables = $S.Variables
        $NumberExec = $S.NumberExec
        $PauseTimeInSec = $S.PauseTimeInSec

        $ScriptFileFullPath = join-path $ScriptFileFolder $ScriptFileName 

        $SqlFileFullName = $ScriptFileFolder + "\" + $ScriptFileName

        $connTimeout = 30; #set to 30 seconds 
        $queryTimeout = 5; # 3 seconds 

        $ReturnValues = @{ }
         
        for ($runNumber = 1; $runNumber -le $NumberExec; $runNumber++ ) {
            $displayMsg = "Pausing for " + $PauseTimeInSec + " Seconds before executing next query in Script File: " + $ScriptFileName
            Write-Host $displayMsg -ForegroundColor Yellow -BackgroundColor Black
            Start-Sleep -s $PauseTimeInSec
        
            $StartDate = (Get-Date)
            $ReturnValues = ExecuteScriptFile -Connection $MySqlConnection -ConnectionTimeout $connTimeout -InputFile $ScriptFileFullPath -QueryTimeout $queryTimeout -Variables $Variables
            $EndDate = (Get-Date)
            $Timespan = (New-TimeSpan -Start $StartDate -End $EndDate)
            $DurationSec = ($Timespan.seconds + ($Timespan.Minutes * 60) + ($Timespan.Hours * 60 * 60))
           
            if ($ReturnValues.Get_Item("Status") -eq 'Success') {
                $Status = $ReturnValues.Get_Item("Status")
                $Message = "  Process Completed for File: " + $SqlFileFullName + " Duration: " + $DurationSec
                Write-Host $Message -ForegroundColor Green -BackgroundColor Black
                #$StatusRow = 0, $ScriptType, $ScriptFileFolder, $ScriptFileName, $Variables, $NumberExec, $PauseTimeInSec, $Status, $DurationSec
                #$StatusRow -join ","  >> $StatusLogFile
                $rows | Add-Member -MemberType NoteProperty -Name "Active" -Value '0' -force	
                $rows | Add-Member -MemberType NoteProperty -Name "ScriptType" -Value $ScriptType -force	
                $rows | Add-Member -MemberType NoteProperty -Name "ScriptFileFolder" -Value $ScriptFileFolder -force	
                $rows | Add-Member -MemberType NoteProperty -Name "ScriptFileName" -Value $ScriptFileName -force	
                $rows | Add-Member -MemberType NoteProperty -Name "Variables" -Value $Variables -force	
                $rows | Add-Member -MemberType NoteProperty -Name "NumberExec" -Value $NumberExec -force	
                $rows | Add-Member -MemberType NoteProperty -Name "PauseTimeInSec" -Value $PauseTimeInSec -force
                $rows | Add-Member -MemberType NoteProperty -Name "Status" -Value $Status -force	
                $rows | Add-Member -MemberType NoteProperty -Name "DurationSec" -Value $DurationSec -force	
                $rows | Export-Csv -Path "$StatusLogFile" -Append -Delimiter "," -NoTypeInformation
                  
            }
            else {
                $ErrorMsg = "  Error running Script for File: " + $FileName + "Error: " + $ReturnValues.Get_Item("Msg") + "Duration: " + $DurationSec + " Seconds"
                Write-Host $ErrorMsg -ForegroundColor Red -BackgroundColor Black
                $Status = "Error: " + $ReturnValues.Get_Item("Msg")
                $Status = $Status.Replace("`r`n", "")
                $Status = '"' + $Status.Replace("`n", "") + '"'
                #$StatusRow = 1, $ScriptType, $ScriptFileFolder, $ScriptFileName, $Variables, $NumberExec, $PauseTimeInSec, $Status, $DurationSec
                #$StatusRow -join ","  >> $StatusLogFile
                $rows | Add-Member -MemberType NoteProperty -Name "Active" -Value '1' -force	
                $rows | Add-Member -MemberType NoteProperty -Name "ScriptType" -Value $ScriptType -force	
                $rows | Add-Member -MemberType NoteProperty -Name "ScriptFileFolder" -Value $ScriptFileFolder -force	
                $rows | Add-Member -MemberType NoteProperty -Name "ScriptFileName" -Value $ScriptFileName -force	
                $rows | Add-Member -MemberType NoteProperty -Name "Variables" -Value $Variables -force	
                $rows | Add-Member -MemberType NoteProperty -Name "NumberExec" -Value $NumberExec -force	
                $rows | Add-Member -MemberType NoteProperty -Name "PauseTimeInSec" -Value $PauseTimeInSec -force
                $rows | Add-Member -MemberType NoteProperty -Name "Status" -Value $Status -force	
                $rows | Add-Member -MemberType NoteProperty -Name "DurationSec" -Value $DurationSec -force	
                $rows | Export-Csv -Path "$StatusLogFile" -Append -Delimiter "," -NoTypeInformation
            }
            
        }

    }

}
#}

#RunParallelExecute

$MySqlConnection.Close()

$FinishTime = Get-Date 
Write-Host "Finished work at " $FinishTime  
Write-Host "Have a great day!" 
