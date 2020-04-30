############################################################################################################
############################################################################################################
#
# Author: Gaiye "Gail" Zhou
# May 2020
# Description: Scripts to run SQL Scripts stored in files 
#   with an option of running stress testing against the SQL Server DB (Or Azure Synapse DB)
#   one user runs selected queries multiple times (configurable)
############################################################################################################
Function GetPassword([SecureString] $securePassword)
{
       $securePassword = Read-Host "Enter Password" -AsSecureString
       $P = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))
       return $P
}


function ExecuteScriptFile { 
	[CmdletBinding()] 
	param( 
        [Parameter(Position = 6, Mandatory = $false)] [System.Data.SqlClient.SqlConnection]$Connection, 
		[Parameter(Position = 7, Mandatory = $false)] [Int32]$ConnectionTimeout = 0, 
        [Parameter(Position = 8, Mandatory = $false)] [string]$InputFile, 
        [Parameter(Position = 6, Mandatory = $false)] [Int32]$QueryTimeout = 0, 
        [Parameter(Position = 10, Mandatory = $false)] [string]$Variables = '',
		[Parameter(Position = 9, Mandatory = $false)] [ValidateSet("DataSet", "DataTable", "DataRow")] [string]$As = "DataSet"
    ) 

    $ReturnValues = @{}
	try {
		if ($InputFile) { 
			$filePath = $(resolve-path $InputFile).path 
			$Query = [System.IO.File]::ReadAllText("$filePath") 
		} 
		if ($Variables) {
		 if ($Variables -match ";") {
			 $splitvar = $Variables.Split(";")
			 foreach ($var in $splitvar) {
					$splitstr = $var.Split("|")
					$search = "(?<![\w\d])" + $splitstr[0] + "(?![\w\d])"
					$replace = $splitstr[1]
					$Query = $Query -replace $search, $replace
			 }
		 }
		 else {
				$splitstr = $Variables.Split("|")
				$search = "(?<![\w\d])" + $splitstr[0] + "(?![\w\d])"
				$replace = $splitstr[1]
				$Query = $Query -replace $search, $replace
		 }
		 		 
		}
     
		#Following EventHandler is used for PRINT and RAISERROR T-SQL statements. Executed when -Verbose parameter specified by caller 
		if ($PSBoundParameters.Verbose) { 
			$Connection.FireInfoMessageEventOnUserErrors = $true 
			$handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] { Write-Verbose "$($_)" } 
			$Connection.add_InfoMessage($handler) 
		} 
     
		#$conn.Open() 
		#$ConnOpen = 'YES'
		$cmd = new-object system.Data.SqlClient.SqlCommand($Query, $Connection) 
		$cmd.CommandTimeout = $QueryTimeout 
		$ds = New-Object system.Data.DataSet 
		$da = New-Object system.Data.SqlClient.SqlDataAdapter($cmd) 

	
		[void]$da.fill($ds) 

		$ReturnValues.add('Status', "Success")
		$ReturnValues.add('Msg', $ErrVar)
	}
	Catch [System.Data.SqlClient.SqlException] { # For SQL exception  
		$Err = $_ 

		$ReturnValues.add('Status', "Error")
		$ReturnValues.add('Msg', $Err)
		
		Write-Verbose "Capture SQL Error" 
		if ($PSBoundParameters.Verbose) { Write-Verbose "SQL Error:  $Err" }  
	} 
	Catch { # For other exception 
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


##################################################################
#
# Main program starts here
#
##################################################################

$ScriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
Set-Location -Path $ScriptPath


# 
# Set up 
# Download the PSExcel Moudle and stored in 
# $psExcelDir = "C:\psExcelDir" 
# Import-Module$psExcelDir"\PSExcel"

$defaultCfgFilePath = $ScriptPath

$displayMsg = " Program starts now. Be prepared to confirm or specify the configuration file. "
Write-Host " "
Write-Host $displayMsg -ForegroundColor Cyan 
Write-Host " "

$cfgFilePath = Read-Host -prompt "Enter the Config File Path or press 'Enter' to accept the default [$($defaultCfgFilePath)]"
if([string]::IsNullOrEmpty($cfgFilePath)) {
    $cfgFilePath = $defaultCfgFilePath
}

$defaultCfgFile = "Stress-Testing-Config.xlsx"

$cfgFile = Read-Host -prompt "Enter the Config File Name or press 'Enter' to accept the default [$($defaultCfgFile)]"
if([string]::IsNullOrEmpty($cfgFile)) {
    $cfgFile = $defaultCfgFile
}
$CfgFileFullPath = join-path $cfgFilePath $cfgFile
if (!(test-path $CfgFileFullPath )) {
    Write-Host "Could not find Config File: $CfgFileFullPath " -ForegroundColor Red
    break 
}


$StatusLogFile = $ScriptPath + "\" + "Stress-Test-SQLScripts-Log.txt"
$HeaderRow = "Active","ScriptType","ScriptFileFolder","ScriptFileName","Variables","NumberExec","PauseTimeInSec","Status","RunDurationSec"
$HeaderRow  -join ","  >> $StatusLogFile

if ((test-path $StatusLogFile)) {
    Write-Host "Replace previous status log file: "$StatusLogFile -ForegroundColor Yellow
    Remove-Item $StatusLogFile -Force
}


# Create an Excel workbook...
$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.WorkBooks.Open($CfgFileFullPath)
$WksServerDbCfg = $Workbook.WorkSheets.Item(1); # Only 1 sheet so this doesn't need to change...
$WksScriptsCfg = $Workbook.WorkSheets.Item(2); # Only 1 sheet so this doesn't need to change...
$StartRow = 2; # ...ignore headers...

#=====================================================================
# Get SQL Server and DB info from 1st sheet of the workbook 
#=====================================================================

#skip the header 
$server =  $WksServerDbCfg.Cells.Item(2,1).Value() # 2nd row, first column
$database =  $WksServerDbCfg.Cells.Item(2,2).value()     # 2nd row, second colum 
$msg = "Server Name: " + $server + " Database Name: " + $database 
Write-Host $msg -ForegroundColor Blue


#========================================
# Get SQL Server Connection 
#=======================================

# Get User Name and Password 
$defaultIntegrated = "Yes"
$Integrated = Read-Host -prompt "Enter 'Yes' or 'No' to connect using integrated Security or press 'Enter' to accept the default [$($defaultIntegrated)]"
if([string]::IsNullOrEmpty($Integrated)) 
{
  $Integrated = $defaultIntegrated
}

if ($Integrated.toUpper() -eq "NO")
{
    Write-Host "Please Enter SQLAUTH Login Information..." -ForegroundColor Yellow
    $UserName = Read-Host -prompt "Enter the User Name"
    if([string]::IsNullOrEmpty($UserName)) 
    {
        Write-Host "A user name must be entered" -ForegroundColor Red
        break
    }
    $Password = GetPassword
    if([string]::IsNullOrEmpty($Password)) 
    {
        Write-Host "A password must be entered." -ForegroundColor Red
        break
    }

}

# default to use Integrated Security 
$MySqlConnection = New-Object System.Data.SqlClient.SqlConnection("Data Source=$server;Integrated Security=SSPI;Initial Catalog=$database;")

if ($Integrated.toUpper() -eq 'NO')
{
    $MySqlConnection = New-Object System.Data.SqlClient.SqlConnection("Data Source=$server;Integrated Security=false;Initial Catalog=$database;User ID=$UserName;Password=$Password")

}


#==================================================
# Read 2nd sheet of the workbook 
#===================================================
$WksScriptsCfg = $Workbook.WorkSheets.Item(2); 
$StartRow = 2; # ...ignore headers...
$rowCounts= $WksScriptsCfg.UsedRange.Rows.Count

$error.Clear()
for ($row = $StartRow; $row -le $rowCounts; $row++)
{
    $error.Clear()
    $ReturnValues = @{}
    #$StartDate=(Get-Date)
    $Active = $WksScriptsCfg.Cells.Item($row, 1).Value();

    $ScriptType = $WksScriptsCfg.Cells.Item($row,2).Value();
    $ScriptFileFolder = $WksScriptsCfg.Cells.Item($row,3).Value();
    $ScriptFileName = $WksScriptsCfg.Cells.Item($row,4).Value();
    $Variables = $WksScriptsCfg.Cells.Item($row,5).Value();
    $NumberExec = $WksScriptsCfg.Cells.Item($row,6).Value();
    $PauseTimeInSec = $WksScriptsCfg.Cells.Item($row,7).Value();

    $ScriptFileFullPath = join-path $ScriptFileFolder $ScriptFileName 

    $SqlFileFullName = $ScriptFileFolder + "\" +  $ScriptFileName

    $connTimeout = 30; #set to 30 seconds 
    $queryTimeout = 5; # 3 seconds 

    if ($Active -eq 1)
    {
        for ($runNumber=1; $runNumber -le $NumberExec; $runNumber++ )
        {
            $displayMsg = "Pausing for " + $PauseTimeInSec + " Seconds before executing next query."
            Write-Host $displayMsg -ForegroundColor Yellow -BackgroundColor Black
            Start-Sleep -s $PauseTimeInSec
    
            $StartDate=(Get-Date)
            $ReturnValues = ExecuteScriptFile -Connection $MySqlConnection -ConnectionTimeout $connTimeout -InputFile $ScriptFileFullPath -QueryTimeout $queryTimeout -Variables $Variables
    
            if($ReturnValues.Get_Item("Status") -eq 'Success')
            {
                $EndDate=(Get-Date)
                $Timespan = (New-TimeSpan -Start $StartDate -End $EndDate)
                $DurationSec = ($Timespan.seconds + ($Timespan.Minutes * 60) + ($Timespan.Hours * 60 * 60))
                $Message = "Process Completed for File: " + $SqlFileFullName + " Duration: " + $DurationSec
                Write-Host $Message -ForegroundColor Green -BackgroundColor Black
                $Status = $ReturnValues.Get_Item("Status")
                $StatusRow = 0,$ScriptType,$ScriptFileFolder,$ScriptFileName,$Variables,$NumberExec,$PauseTimeInSec,$Status,$DurationSec
                $StatusRow  -join ","  >> $StatusLogFile
               }
            else
        
            {
                 $EndDate=(Get-Date)
                 $Timespan = (New-TimeSpan -Start $StartDate -End $EndDate)
                 $DurationSec = ($Timespan.seconds + ($Timespan.Minutes * 60) + ($Timespan.Hours * 60 * 60))
                 $ErrorMsg = "Error running Script for File: " + $FileName + "Error: " + $ReturnValues.Get_Item("Msg") + "Duration: " + $DurationSec + " Seconds"
                 Write-Host $ErrorMsg -ForegroundColor Red -BackgroundColor Black
                 $Status = "Error: " + $ReturnValues.Get_Item("Msg")
                 $Status = $Status.Replace("`r`n", "")
                 $Status = '"' + $Status.Replace("`n", "") + '"'
                 $StatusRow = 0,$ScriptType,$ScriptFileFolder,$ScriptFileName,$Variables,$NumberExec,$PauseTimeInSec,$Status,$DurationSec
                 $StatusRow  -join ","  >> $StatusLogFile
            }
        
        }
  
    }
  
}

$Excel.Workbooks.Close()
$Excel.Quit() 

$MySqlConnection.Close()

$FinishTime = Get-Date 
Write-Host "Finished work at " $FinishTime  
Write-Host "Have a great day!" 
