############################################################################################################
############################################################################################################
#
# Author: Gaiye "Gail" Zhou
# May 2020
# Description: It will run SQL Scripts stored in .sql files in specified folder. 
# The program will promot users for below information
#   (1) Folder Name where SQL Scripts are stored
#   (2) Full Qualified SQL Server Name or Azure Synapse SQL Server Name (Server.database.windows.net)
#   (3) Using Integrated Security or Not 
#   (4) SQL Authenfication User Name and Password (if not using Integrated Security)
# The program will produce a log file for status of runing the SQL Scripts. 
# The log file name is RunSqlFilesInFolder_Log.csv   (If you name this script RunSqlFilesInFolder.ps1)
#  
############################################################################################################


Function GetPassword([SecureString] $securePassword) {
    $securePassword = Read-Host "Enter Password" -AsSecureString
    $P = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))
    return $P
}

Function ExecuteScriptFile { 
    [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [System.Data.SqlClient.SqlConnection]$Connection, 
        [Parameter(Position = 2, Mandatory = $false)] [string]$InputFile, 
        [Parameter(Position = 3, Mandatory = $false)] [Int32]$QueryTimeout = 30
    ) 

    $ReturnValues = @{ }
    try {
        $filePath = $(resolve-path $InputFile).path 
        $Query = [System.IO.File]::ReadAllText("$filePath") 

        $cmd = new-object system.Data.SqlClient.SqlCommand($Query, $Connection) 
        $cmd.CommandTimeout = $QueryTimeout 
        $ds = New-Object system.Data.DataSet 
        $da = New-Object system.Data.SqlClient.SqlDataAdapter($cmd)

        [void]$da.fill($ds) 

        $ReturnValues.add('Status', "Success")
        $ReturnValues.add('Msg', "")

    }
    Catch [System.Data.SqlClient.SqlException] {
        $Err = $_ 
        $ReturnValues.add('Status', "Error")
        $ReturnValues.add('Msg', $Err)
    } 
    Catch {
        $Err = $_ 
        $ReturnValues.add('Status', "Error")
        $ReturnValues.add('Msg', $Err)
    }  
    Finally { 
        $cmd.Dispose()
        $ds.Dispose()
        $da.Dispose()
    }
    return $ReturnValues
	 
} 

###############################################################
# Main Program Starts Here 
###############################################################

$MySqlScriptFilePath = 'C:\Z_Tests\SQLScripts'
$MySqlServer = '.\GailzSqlSvr2017'
$MyDatabase = 'AdventureWorksDW2017'
$DefaultIntegrated = 'YES'

$SqlScriptFilePath = Read-Host -prompt "Enter the Folder of the SQL Script Files or press 'Enter' to accept the default [$($MySqlScriptFilePath)]"
if ([string]::IsNullOrEmpty($SqlScriptFilePath)) {
    $SqlScriptFilePath = $MySqlScriptFilePath
}

$server = Read-Host -prompt "Enter the SQL Server Name  or press 'Enter' to accept the default [$($MySqlServer)]"
if ([string]::IsNullOrEmpty($server)) {
    $server = $MySqlServer
}

$database = Read-Host -prompt "Enter the Database Name  or press 'Enter' to accept the default [$($MyDatabase)]"
if ([string]::IsNullOrEmpty($database)) {
    $database = $MyDatabase
}

$Integrated = Read-Host -prompt "Enter 'Yes' or 'No' to connect using integrated security or press 'Enter' to accept the default [$($DefaultIntegrated)]"
if ([string]::IsNullOrEmpty($Integrated)) {
    $Integrated = $DefaultIntegrated
}

if ($Integrated.toUpper() -eq "NO") {
    Write-Host "Please Enter SQLAUTH Login Information..." -ForegroundColor Yellow
    $UserName = Read-Host -prompt "Enter the User Name: "
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

$connectionTimeOut = '30' # in seconds 
$MySqlConnection = New-Object System.Data.SqlClient.SqlConnection("Data Source=$server;Integrated Security=SSPI;Initial Catalog=$database;Connect Timeout=$connectionTimeOut")

if ($Integrated.toUpper() -eq 'NO') {
    $MySqlConnection = New-Object System.Data.SqlClient.SqlConnection("Data Source=$server;Integrated Security=false;Initial Catalog=$database;User ID=$UserName;Password=$Password;Connect Timeout=$connectionTimeOut")    
}

$MySqlConnection.open()


$ScriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
Set-Location -Path $ScriptPath

$MyScriptFile = $MyInvocation.MyCommand.Path 

$MyScriptFileWoExt = [System.IO.Path]::GetFileNameWithoutExtension($MyScriptFile)
$LogFileFullPath = $MyScriptFileWoExt + "_Log.csv"

if ((test-path $LogFileFullPath)) {
    Write-Host "Replace previous log file: "$LogFileFullPath -ForegroundColor Magenta
    Remove-Item $LogFileFullPath -Force
}


$HeaderRow = "SqlScriptFile", "DurationSec", "Status", "Message"
$HeaderRow -join ","  >> $LogFileFullPath

$myQueryTimeOut = '30' 
foreach ($f in Get-ChildItem -path $SqlScriptFilePath  -Filter *.sql) {
    #Write-Host "File Name: " $f.FullName.ToString()	
    $ReturnValues = @{ }
    
    $SqlScriptFileName = $f.FullName.ToString()	

    # Run the Script in $SqlScriptFilename 
    $StartDate = (Get-Date)
    $ReturnValues  = ExecuteScriptFile -Connection $MySqlConnection -InputFile $SqlScriptFileName -QueryTimeout $myQueryTimeOut 
    $Status = $ReturnValues.Get_Item("Status")
    $Message = $ReturnValues.Get_Item("Msg")
    $EndDate = (Get-Date)
    $Timespan = (New-TimeSpan -Start $StartDate -End $EndDate)
    $DurationSec = ($Timespan.Seconds + ($Timespan.Minutes * 60) + ($Timespan.Hours * 60 * 60))

    if ($ReturnValues.Get_Item("Status") -eq 'Success') {
        $DisplayMessage = "  Process Completed for File: " + $SqlScriptFileName + " Duration: " + $DurationSec + " Seconds."
        Write-Host $DisplayMessage -ForegroundColor Green -BackgroundColor Black
    }
    else {
        $Message = "Error: " + $Message
        $Message = $Message.Replace("`r`n", "")
        $Message = '"' + $Message.Replace("`n", "") + '"'
        $DisplayMessage = "  Error Processing File: " + $SqlScriptFileName  + ". Error: " + $Message 
        Write-Host $DisplayMessage -ForegroundColor Red -BackgroundColor Black
    }
    $ReturnValues.Clear()
 
    $dataRow = $SqlScriptFileName, $DurationSec, $Status, $Message
    $dataRow -join ","  >> $LogFileFullPath
}

$MySqlConnection.close()

