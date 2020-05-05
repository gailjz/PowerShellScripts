############################################################################################################
############################################################################################################
#
# Author: Gaiye "Gail" Zhou
# May 2020
# Description: Scripts to run SQL Scripts stored in files 
#
############################################################################################################
Function ExecuteScriptFile { 
    [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [System.Data.SqlClient.SqlConnection]$Connection, 
        [Parameter(Position = 2, Mandatory = $false)] [Int32]$ConnectionTimeout = 0, 
        [Parameter(Position = 3, Mandatory = $false)] [string]$InputFile, 
        [Parameter(Position = 4, Mandatory = $false)] [Int32]$QueryTimeout = 0, 
        [Parameter(Position = 5, Mandatory = $false)] [string]$Variables = '',
        [Parameter(Position = 6, Mandatory = $false)] [ValidateSet("DataSet", "DataTable", "DataRow")] [string]$As = "DataSet"
    ) 

    $ReturnValues = @{ }
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
    
        $cmd = new-object system.Data.SqlClient.SqlCommand($Query, $Connection) 
        $cmd.CommandTimeout = $QueryTimeout 
        $ds = New-Object system.Data.DataSet 
        $da = New-Object system.Data.SqlClient.SqlDataAdapter($cmd) 

	
        [void]$da.fill($ds) 
        #$da.fill($ds) 

        $ReturnValues.add('Status', "Success")
        $ReturnValues.add('Msg', $ErrVar)

        #PrintDataSetRows -dataSet $ds
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


Function ProcessConfigAndRunScript
{
    [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [string]$csvFileFullPath = '',
        [Parameter(Position = 2, Mandatory = $false)] [string]$statusLogFileFullPath = '',
        [Parameter(Position = 3, Mandatory = $false)] [System.Data.SqlClient.SqlConnection]$Connection
    ) 
    $scriptsCfgCsv = $scriptsCfgCsv = Import-Csv $csvFileFullPath

    ForEach ($S in $scriptsCfgCsv ) {
        $StartDate = (Get-Date)
        $Active = $S.Active
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
            $rows = New-Object PSObject 
            for ($runNumber = 1; $runNumber -le $NumberExec; $runNumber++ ) {
                $displayMsg = "Pausing for " + $PauseTimeInSec + " Seconds before executing next query in Script File: " + $ScriptFileName
                Write-Host $displayMsg -ForegroundColor Yellow -BackgroundColor Black
                Start-Sleep -s $PauseTimeInSec
            
                $StartDate = (Get-Date)
                $ReturnValues = ExecuteScriptFile -Connection $MySqlConnection -ConnectionTimeout $connTimeout -InputFile $ScriptFileFullPath -QueryTimeout $queryTimeout -Variables $Variables
                $EndDate = (Get-Date)
                $Timespan = (New-TimeSpan -Start $StartDate -End $EndDate)
                $DurationSec = ($Timespan.Seconds + ($Timespan.Minutes * 60) + ($Timespan.Hours * 60 * 60))
               
                if ($ReturnValues.Get_Item("Status") -eq 'Success') {
                    $Status = $ReturnValues.Get_Item("Status")
                    $Message = "  Process Completed for File: " + $SqlFileFullName + " Duration: " + $DurationSec + " Seconds."
                    Write-Host $Message -ForegroundColor Green -BackgroundColor Black
                    $rows | Add-Member -MemberType NoteProperty -Name "Active" -Value '0' -force    
                }
                else {
                    $ErrorMsg = "  Error running Script for File: " + $FileName + "Error: " + $ReturnValues.Get_Item("Msg") + "Duration: " + $DurationSec + " Seconds"
                    Write-Host $ErrorMsg -ForegroundColor Red -BackgroundColor Black
                    $Status = "Error: " + $ReturnValues.Get_Item("Msg")
                    $Status = $Status.Replace("`r`n", "")
                    $Status = '"' + $Status.Replace("`n", "") + '"'
                    $rows | Add-Member -MemberType NoteProperty -Name "Active" -Value '1' -force	
                }
                $rows | Add-Member -MemberType NoteProperty -Name "ScriptType" -Value $ScriptType -force	
                $rows | Add-Member -MemberType NoteProperty -Name "ScriptFileFolder" -Value $ScriptFileFolder -force	
                $rows | Add-Member -MemberType NoteProperty -Name "ScriptFileName" -Value $ScriptFileName -force	
                $rows | Add-Member -MemberType NoteProperty -Name "Variables" -Value $Variables -force	
                $rows | Add-Member -MemberType NoteProperty -Name "NumberExec" -Value $NumberExec -force	
                $rows | Add-Member -MemberType NoteProperty -Name "PauseTimeInSec" -Value $PauseTimeInSec -force
                $rows | Add-Member -MemberType NoteProperty -Name "Status" -Value $Status -force	
                $rows | Add-Member -MemberType NoteProperty -Name "DurationSec" -Value $DurationSec -force	
                $rows | Export-Csv -Path "$statusLogFileFullPath" -Append -Delimiter "," -NoTypeInformation
    
                
            }
    
        }
    
    }



}
