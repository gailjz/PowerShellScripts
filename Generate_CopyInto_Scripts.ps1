############################################################################################################
############################################################################################################
#
# Author: Gaiye "Gail" Zhou
# April 2020
# Description: Generate COPY Into T-SQL Scripts to migrate date into Synapse 
#
############################################################################################################

function CreateCopyIntoScripts(
$StorageType="BLOB",
$CredentialType="Managed Identity",
$Credential="(IDENTITY= 'Managed Identity')",
$AccountName="BlobAccountName",
$Container="ContainerName",
$FileType="CSV",
$FieldQuote,
$FieldTerminator,
$RowTerminator,
$Encoding,
$FirstRow="2",
$TruncateTable="YES", 
$Folder = "DbName",
$SchemaName,
$TableName,
$SqlFilePath="C:\migratemaster\tsql\asa_objects\CopyInto",
$SqlFileName="Test.sql"
)
{
	if (!(test-path $SqlFilePath))
	{
    New-item "$SqlFilePath" -ItemType Dir | Out-Null
    #New-item "$SqlFilePath" -ItemType File | Out-Null
  }

  $SqlFileFullPath = join-path $SqlFilePath $SqlFileName

	if ((test-path $SqlFileFullPath))
	{
    Write-Host "Replace output file: "$SqlFileFullPath -ForegroundColor Red
    Remove-Item $SqlFileFullPath -Force
  }
    
  $Storage = $StorageType.toUpper()
  $Location = "'https://accountname.blob.core.windows.net/import/DimAccount/'"
  if ($Storage -eq 'BLOB')
  {
      $Location = "https://" + $AccountName + ".blob.core.windows.net/" + $Container + "/" + $Folder + "/"
  }
  elseif ($Storage -eq 'ADLS')
  { 
    Write-Host "TODO: Later"
  }

  if ($TruncateTable.toUpper() -eq 'YES')
  {
    " " >> $SqlFileFullPath
    "Truncate Table " + $SchemaName + "."+ $TableName >> $SqlFileFullPath
    " " >> $SqlFileFullPath
  }
  "COPY INTO " + $SchemaName + "."+ $TableName >> $SqlFileFullPath
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
}

  
#==========================================================================================================
# Main Program Starts here 
# Generate COPY Into T-SQL Scripts to migrate date into Synapse 
#==========================================================================================================


$ScriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
Set-Location -Path $ScriptPath


$defaultCfgFilePath = "C:\migratemaster\input"

$cfgFilePath = Read-Host -prompt "Enter the Config File Path or press 'Enter' to accept the default [$($defaultCfgFilePath)]"
if([string]::IsNullOrEmpty($cfgFilePath)) 
{
  $cfgFilePath = $defaultCfgFilePath
}


#$defaultCfgFile = "CopyIntoConfig_mi.json"
$defaultCfgFile = "CopyIntoConfig_mi.json"
$cfgFile = Read-Host -prompt "Enter the Config File Name or press 'Enter' to accept the default [$($defaultCfgFile)]"
if([string]::IsNullOrEmpty($cfgFile)) 
{
  $cfgFile = $defaultCfgFile
}

#$BCPDriverFileFullPath = join-path $ScriptPath $BCPDriverFile
$CfgFileFullPath = join-path $cfgFilePath $cfgFile

if (!(test-path $CfgFileFullPath ))
{
  Write-Host "Could not find Config File: $CfgFileFullPath " -ForegroundColor Red
  break 
}

Write-Host $CfgFileFullPath

#$PowerShellObject=Get-Content -Path settings.json | ConvertFrom-Json


#$BaseJSON = (Get-Content $CfgFileFullPath) -join "`n" | ConvertFrom-Json
$JsonConfig = Get-Content -Path $CfgFileFullPath | ConvertFrom-Json

# Read sections of JSON file 
#$StorageConfig = ($BaseJSON | Select-Object StorageConfig).StorageConfig 
#$FileConfig = ($BaseJSON | Select-Object FileConfig).FileConfig 
#$InputFilePath = ($BaseJSON | Select-Object InputFilePath).InputFilePath 

Write-Host $JsonConfig

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


$Folder = "DbName/Tables"
$SchemaName = "deploy"
$TableName = "DimAccount"
$SqlFilePath = "C:\migratemaster\tsql\asa_objects\CopyInto"
$SqlFileName= "Test.sql"

CreateCopyIntoScripts -StorageType $StorageType -CredentialType $CredentialType -Credential $Credential -AccountName $AccountName -Container $Container `
  -FileType $FileType -FieldQuote $FieldQuote -FieldTerminator $FieldTerminator -RowTerminator $RowTerminator  -Encoding $Encoding -FirstRow $FirstRow `
  -TruncateTable $TruncateTable -Folder $Folder -SchemaName $SchemaName -TableName $TableName -SqlFilePath $SqlFilePath -SqlFileName $SqlFileName
