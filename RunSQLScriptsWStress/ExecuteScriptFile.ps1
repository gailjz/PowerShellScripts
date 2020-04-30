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
