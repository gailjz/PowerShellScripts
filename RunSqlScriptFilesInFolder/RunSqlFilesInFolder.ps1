############################################################################################################
############################################################################################################
#
# Author: Gaiye "Gail" Zhou
# May 2020
# Description: Scripts to run SQL Scripts stored in files 
#  
############################################################################################################


Function GetPassword([SecureString] $securePassword) {
    $securePassword = Read-Host "Enter Password" -AsSecureString
    $P = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))
    return $P
}

Function ExecuteScriptFileLogResult { 
    [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [System.Data.SqlClient.SqlConnection]$Connection, 
        [Parameter(Position = 2, Mandatory = $false)] [string]$InputFile, 
        [Parameter(Position = 3, Mandatory = $false)] [Int32]$QueryTimeout = 30
    ) 

    $ReturnValues = @{ }
    try {
        if ($InputFile) { 
            $filePath = $(resolve-path $InputFile).path 
            $Query = [System.IO.File]::ReadAllText("$filePath") 
        } 
     
        #Following EventHandler is used for PRINT and RAISERROR T-SQL statements. Executed when -Verbose parameter specified by caller 
        if ($PSBoundParameters.Verbose) { 
            $Connection.FireInfoMessageEventOnUserErrors = $true 
            $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] { Write-Verbose "$($_)" } 
            $Connection.add_InfoMessage($handler) 
        } 
    
        $cmd = new-object system.Data.SqlClient.SqlCommand($Query, $Connection) 
        $cmd.CommandTimeout = $QueryTimeout 
        $ds = New-Object system.Data.DataSet 
        $da = New-Object system.Data.SqlClient.SqlDataAdapter($cmd)

        [void]$da.fill($ds) 
        #$da.fill($ds) 

        $ReturnValues.add('Status', "Success")
        $ReturnValues.add('Msg', "Done")

    }
    Catch [System.Data.SqlClient.SqlException] {
        # For SQL exception  
        $Err = $_ 

        $ReturnValues.add('Status', "Error")
        $ReturnValues.add('Msg', $Err)
		
        Write-Verbose "Capture SQL Error" 
        if ($PSBoundParameters.Verbose) { Write-Verbose "SQL Error:  $Err" }  
    } 
    Catch {
        # For other exception 
        #	Write-Verbose "Capture Other Error"   

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


$Integrated = Read-Host -prompt "Enter 'Yes' or 'No' to connect using integrated Security or press 'Enter' to accept the default [$($DefaultIntegrated)]"
if ([string]::IsNullOrEmpty($Integrated)) {
    $Integrated = $DefaultIntegrated
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

$connectionTimeOut = '30' # in seconds 

#$MySqlConnection = New-Object System.Data.SqlClient.SqlConnection("Data Source=$server;Integrated Security=SSPI;Initial Catalog=$database;)
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


$HeaderRow = "ScriptFile", "Status", "Message", "DurationSec"
$HeaderRow -join ","  >> $LogFileFullPath

$myQueryTimeOut = '30' 
foreach ($f in Get-ChildItem -path $SqlScriptFilePath  -Filter *.sql) {
    #Write-Host "File Name: " $f.FullName.ToString()	
    $ReturnValues = @{ }
    
    $SqlScriptFileName = $f.FullName.ToString()	
    $StartDate = (Get-Date)
    $ReturnValues  = ExecuteScriptFileLogResult -Connection $MySqlConnection -InputFile $SqlScriptFileName -QueryTimeout $myQueryTimeOut 
    $Status = $ReturnValues.Get_Item("Status")
    $Message = $ReturnValues.Get_Item("Msg")
    $EndDate = (Get-Date)
    $Timespan = (New-TimeSpan -Start $StartDate -End $EndDate)
    $DurationSec = ($Timespan.Seconds + ($Timespan.Minutes * 60) + ($Timespan.Hours * 60 * 60))
    $rows = New-Object PSObject 
    if ($ReturnValues.Get_Item("Status") -eq 'Success') {
        $DisplayMessage = "  Process Completed for File: " + $SqlScriptFileName + " Duration: " + $DurationSec + " Seconds."
        Write-Host $DisplayMessage -ForegroundColor Green -BackgroundColor Black
    }
    else {
        $ErrorMsg = "  Error running Script for File: " + $SqlScriptFileName  + "Error: " + $ReturnValues.Get_Item("Msg") + "Duration: " + $DurationSec + " Seconds"
        Write-Host $ErrorMsg -ForegroundColor Red -BackgroundColor Black
        $Status = "Error"
        $Message = "Error Message: " + $ReturnValues.Get_Item("Msg")
        $Message = $Message.Replace("`r`n", "")
        $Message = '"' + $Message.Replace("`n", "") + '"'
    }

    $DataRow = $SqlScriptFileName, $Status, $Message, $DurationSec
    $dataRow -join ","  >> $LogFileFullPath

    $ReturnValues.clear()
}

$MySqlConnection.close()
