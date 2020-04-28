############################################################################################################
############################################################################################################
#
# Author: Gaiye "Gail" Zhou
# April 2020
# Description: Generate COPY Into T-SQL Scripts to migrate date into Synapse 
#   Azure Storage Type: BLOB or DFS (Works with ADLS Blob Storage as well) 
#   File Format: CSV or Parquet (Preferred) 
#   Authorization: Account Key or Managed Identity (Preferred) 
#   Storage Type, File Format, and Authorization are all configurable 
# Example Configuration Files
#   (1) CopyIntoConfig_blob_key_csv.json      for BLOB and CSV File, Account Key Authorization 
#   (2) CopyIntoConfig_blob_mi_csv.json       for Blob and CSV File, Managed Identity Authorization 
#   (3) CopyIntoConfig_blob_mi_parquet.json   for Blob and Parquet File, Managed Identity Authorization 
#   (4) CopyIntoConfig_dfs_mi_csv.json        for DFS and CSV File, Managed Identity Authorization 
#   (5) CopyIntoConfig_dfs_mi_parquet.json    for DFS and Parquet File, Managed Identity Authorization 
############################################################################################################

function CreateCopyIntoScripts(
$StorageType="BLOB",
$CredentialType="Managed Identity",
$Credential="(IDENTITY= 'Managed Identity')",
$AccountName="BlobAccountName",
$Container="ContainerName",
$FileType="CSV",
$FieldQuote = '"',
$FieldTerminator="^|^",
$RowTerminator="0x0A",
$Encoding="UTF8",
$FirstRow="2",
$TruncateTable="YES", 
$StorageFolder ="DbName/Tables/Hello",
$SchemaName="dbo",
$TableName ="TableName",
$AsaSchema = "deploy",
$SqlFilePath="C:\migratemaster\tsql\asa_objects\CopyInto",
$SqlFileName="Test.sql"
)
{
	if (!(test-path $SqlFilePath)) {
        New-item "$SqlFilePath" -ItemType Dir | Out-Null
        #New-item "$SqlFilePath" -ItemType File | Out-Null
    }

    $SqlFileFullPath = join-path $SqlFilePath $SqlFileName

	if ((test-path $SqlFileFullPath)) {
        Write-Host "Replace output file: "$SqlFileFullPath -ForegroundColor Yellow
        Remove-Item $SqlFileFullPath -Force
    }

  
    #============================================================
    # Storage Types: BLOB or DFS 
    #============================================================
    
    $Storage = $StorageType.toUpper()
    # Defulat Location 
    $Location = "'https://accountname.blob.core.windows.net/import/DimAccount/'" 

    if ($Storage -eq 'BLOB') {
        $Location = "https://" + $AccountName + ".blob.core.windows.net/" + $Container + "/" + $StorageFolder + "/"
    } elseif ($Storage -eq 'DFS') { 
        $Location = "abfss://" + $AccountName + ".dfs.core.windows.net/" + $Container + "/" + $StorageFolder + "/"
    } else {
        Write-Host "Unknown Storate Type" $Storage -ForegroundColor Red
        Write-Host "Accepted Value: BLOB or DFS" -ForegroundColor Red
        return -1 
    } 

    if ($TruncateTable.toUpper() -eq 'YES') {
        " " >> $SqlFileFullPath
        "Truncate Table " + $AsaSchema + "."+ $TableName >> $SqlFileFullPath
        " " >> $SqlFileFullPath
    }
  

    #============================================================
    # File Types: CSV or  PARQUET
    #============================================================
    if ($FileType.toUpper() -eq 'CSV') {
        "COPY INTO " + $AsaSchema + "."+ $TableName >> $SqlFileFullPath
        "FROM " + "'" + $Location + "'" >> $SqlFileFullPath
        "WITH (" >> $SqlFileFullPath
        "  FILE_TYPE = " + "'" + $FileType + "'" +"," >> $SqlFileFullPath
        "  CREDENTIAL = " + $Credential +"," >> $SqlFileFullPath
        "  FIELDQUOTE = " + "'" + $FieldQuote + "'" +"," >> $SqlFileFullPath
        "  FIELDTERMINATOR = " + "'" + $FieldTerminator + "'" +"," >> $SqlFileFullPath
        "  ROWTERMINATOR = " + "'" + $RowTerminator + "'" +"," >> $SqlFileFullPath
        "  ENCODING = " + "'" + $Encoding + "'" +"," >> $SqlFileFullPath
        "  FIRSTROW = " + $FirstRow >> $SqlFileFullPath
        ") " >> $SqlFileFullPath
    } elseif ($FileType.toUpper() -eq 'PARQUET') {
        "COPY INTO " + $AsaSchema + "."+ $TableName >> $SqlFileFullPath
        "FROM " + "'" + $Location + "'" >> $SqlFileFullPath
        "WITH (" >> $SqlFileFullPath
        "  FILE_TYPE = " + "'" + $FileType + "'" +"," >> $SqlFileFullPath
        "  CREDENTIAL = " + $Credential >> $SqlFileFullPath
        ") " >> $SqlFileFullPath
    } else {
        Write-Host "Unknown File Type" $FileType -ForegroundColor Red
        Write-Host "Accepted Value: CSV or PARQUET " -ForegroundColor Red
        return -1 
    } 

    return 0
}

  
#==========================================================================================================
# Main Program Starts here 
# Generate COPY Into T-SQL Scripts to migrate date into Synapse 
#==========================================================================================================


$ScriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
Set-Location -Path $ScriptPath


$defaultCfgFilePath = "C:\migratemaster\input"

$cfgFilePath = Read-Host -prompt "Enter the Config File Path or press 'Enter' to accept the default [$($defaultCfgFilePath)]"
if([string]::IsNullOrEmpty($cfgFilePath)) {
    $cfgFilePath = $defaultCfgFilePath
}


#$defaultCfgFile = "CopyIntoConfig_blob_key_csv.json"
#$defaultCfgFile = "CopyIntoConfig_blob_mi_csv.json"
$defaultCfgFile = "CopyIntoConfig_blob_mi_parquet.json"

#$defaultCfgFile = "CopyIntoConfig_dfs_mi_csv.json"
#$defaultCfgFile = "CopyIntoConfig_dfs_mi_parquet.json"

$cfgFile = Read-Host -prompt "Enter the Config File Name or press 'Enter' to accept the default [$($defaultCfgFile)]"
if([string]::IsNullOrEmpty($cfgFile)) {
    $cfgFile = $defaultCfgFile
}
$CfgFileFullPath = join-path $cfgFilePath $cfgFile
if (!(test-path $CfgFileFullPath )) {
    Write-Host "Could not find Config File: $CfgFileFullPath " -ForegroundColor Red
    break 
}

# CSV File
$defaultTablesCfgFile = "CopyIntoTables_CFG.csv"
$tablesCfgFile = Read-Host -prompt "Enter the COPY INTO Tables Config Name or press 'Enter' to accept the default [$($defaultTablesCfgFile)]"
if([string]::IsNullOrEmpty($tablesCfgFile)) {
    $tablesCfgFile = $defaultTablesCfgFile
}
$tablesCfgFileFullPath = join-path $cfgFilePath $tablesCfgFile
if (!(test-path $tablesCfgFileFullPath )) {
    Write-Host "Could not find Config File: $tablesCfgFileFullPath " -ForegroundColor Red
    break 
}

$csvTablesCfgFile = Import-Csv $tablesCfgFileFullPath

$JsonConfig = Get-Content -Path $CfgFileFullPath | ConvertFrom-Json 
#Write-Host $JsonConfig

$StorageType = $JsonConfig.StorateType
$CredentialType = $JsonConfig.CredentialType
$Credential = $JsonConfig.Credential
$AccountName = $JsonConfig.AccountName
$Container = $JsonConfig.Container
$FileType = $JsonConfig.FileType
$FieldQuote = $JsonConfig.FieldQuote
$FieldTerminator = $JsonConfig.FieldTerminator
$RowTerminator = $JsonConfig.RowTerminator
$Encoding = $JsonConfig.Encoding
$FirstRow = $JsonConfig.FirstRow
$TruncateTable = $JsonConfig.TruncateTable

ForEach ($csvItem in $csvTablesCfgFile) {
    $Active = $csvItem.Active
    If ($Active -eq "1") {
      $DatabaseName = $csvItem.DatabaseName
      $SchemaName = $csvItem.SchemaName
      $TableName = $csvItem.TableName
      $AsaSchema = $csvItem.AsaSchema
      $SqlFilePath = $csvItem.SqlFilePath
      $SqlFileName =  "Copy_" + $AsaSchema  + "_" + $TableName + ".sql"

      $StorageFolder =  $DatabaseName + "/" + $SchemaName + "_" + $TableName 

      $returnValue = -1; 
     
      $returnValue = CreateCopyIntoScripts -StorageType $StorageType -CredentialType $CredentialType -Credential $Credential -AccountName $AccountName -Container $Container `
      -FileType $FileType -FieldQuote $FieldQuote -FieldTerminator $FieldTerminator -RowTerminator $RowTerminator  -Encoding $Encoding -FirstRow $FirstRow `
      -TruncateTable $TruncateTable -StorageFolder $StorageFolder -SchemaName $SchemaName -TableName $TableName -AsaSchema $AsaSchema -SqlFilePath $SqlFilePath -SqlFileName $SqlFileName

      if ($returnValue -ne 0) {
        Write-Host "Something went wrong. Please check program or input " -ForegroundColor Red
      }


    }

  }
