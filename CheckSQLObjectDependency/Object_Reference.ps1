

############################################################################################################
############################################################################################################
#
# Author: Gaiye "Gail" Zhou
# March 2020
# 
############################################################################################################
# Description
# File Processing Utilites 
#
############################################################################################################


# Input Configration File that lists all the key OBjects 
#defaultConfigFileName = "Base_Object_List_Sample.csv"
$configFileName = Read-Host -prompt "Enter configuration file name in your PowerShell Script Dir or 'Enter' to accept the default: [$($defaultConfigFileName)]"
if([String]::IsNullOrEmpty($configFileName)) {$configFileName = $defaultConfigFileName} 


# Input Query Results CSV file that captured Object Dependencies 
# This is CSV produced by running Check_Dependencies.sql 
$defaultObjectDepFileName = "EDW_Object_dependencies.csv"

$ObjectDepFileName = Read-Host -prompt "Enter Object Dependency file name in your PowerShell Script Dir or 'Enter' to accept the default: [$($defaultObjectDepFileName)]"
if([String]::IsNullOrEmpty($ObjectDepFileName)) {$ObjectDepFileName = $defaultObjectDepFileName} 

# Consturct output file name using input file as prefix 

$nameLen = $configFileName.length
$defaultOutFileName  = $configFileName.Substring(0,$nameLen - 4) + "_Output" + ".csv" 


$outFileName = Read-Host -prompt "Enter output file name or 'Enter' to accept the default: [$($defaultOutFileName)]"
if([String]::IsNullOrEmpty($outFileName)) {$outFileName= $defaultOutFileName} 


$ScriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent

$configFilePath = $ScriptPath + "\" + $configFileName
$objDepFilePath = $ScriptPath + "\" + $ObjectDepFileName
$outputFilePath = $ScriptPath + "\" + $outFileName 


if (Test-Path $outputFilePath)
{ 
    Write-Host "Replace output file: "$outputFilePath -ForegroundColor Red
    #Remove-Item $outputFilePath -Confirm
    Remove-Item $outputFilePath -Force
} 

$startTime = Get-Date 

$configFileCsv = Import-Csv $configFilePath
$ObjectDepFileCsv  = Import-Csv $objDepFilePath


ForEach ($csvItem in $configFileCsv) 
{

    $Active = $csvItem.Active

    if ($Active -eq '1')
    {
        $BaseDatabaseName = $csvItem.Database
        $BaseSchemaName = $csvItem.Schema
        $BaseObjectType = $csvItem.ObjectType 
        $BaseObjectName = $csvItem.ObjectName 
        $BaseArea= $csvItem.Area 
        $BaseDescription = $csvItem.Description 

        if ([string]::IsNullOrEmpty($BaseObjectName)) { continue; }
        ForEach ($obj in $ObjectDepFileCsv)
        {

            
            $refSchemaName = $obj.referencing_schema
            $refObjeName = $obj.referencing_entity_name
            $refObjTwoPartName = $refSchemaName + "." + $refObjeName 

            $depDbName = $obj.referenced_database_name.ToUpper() # Change DB to Upper
            $depObjType = $obj.referencing_class_desc  # this name is confusing. 
            $depSchemaName = $obj.referenced_schema_name.ToLower()
            $depObjName = $obj.referenced_entity_name 
            $depObjThreePartName =  $depDbName + "." + $depSchemaName + "." + $depObjName

        
            if ([string]::IsNullOrEmpty($refObjeName)) { continue; }

            # This one assumes the person writes down accurate information on Base Schema 
            #if ( ($BaseObjectName.Trim().ToUpper() -eq $refObjeName.Trim().ToUpper()) -and ($BaseSchemaName.Trim().ToUpper() -eq $refSchemaName.Trim().ToUpper()) )
            # -- no schema comparison 
            if ( ($BaseObjectName.Trim().ToUpper() -eq $refObjeName.Trim().ToUpper()) )
            {
     
                $row = New-Object PSObject 	
                # Key information from Config File copied 
                $row | Add-Member -MemberType NoteProperty -Name "BaseDatabase" -Value $BaseDatabaseName -force
                $row | Add-Member -MemberType NoteProperty -Name "BaseSchemaName" -Value $BaseSchemaName -force
                $row | Add-Member -MemberType NoteProperty -Name "BaseObjectType" -Value $BaseObjectType -force
                $row | Add-Member -MemberType NoteProperty -Name "BaseObjectName" -Value $BaseObjectName -force
                $row | Add-Member -MemberType NoteProperty -Name "RefObjectTwoPartName" -Value $refObjTwoPartName -force
           
                # attributes in the dependency objects file 
                $row | Add-Member -MemberType NoteProperty -Name "DepDbName" -Value $depDbName -force
                $row | Add-Member -MemberType NoteProperty -Name "DepObjType" -Value $depObjType -force
                $row | Add-Member -MemberType NoteProperty -Name "DepSchemaName" -Value $depSchemaName -force  
                $row | Add-Member -MemberType NoteProperty -Name "DepObjName" -Value $depObjName -force  
                $row | Add-Member -MemberType NoteProperty -Name "DepObjThreePartName" -Value $depObjThreePartName -force  
                
                # Add Additional Inforamtion from Config file 
                $row | Add-Member -MemberType NoteProperty -Name "BaseArea" -Value $BaseArea -force
                $row | Add-Member -MemberType NoteProperty -Name "BaseDescription" -Value $BaseDescription -force


                #$rows | Export-Csv -Path "$FileName" -Append -Delimiter "," -NoTypeInformation 
                $row | Export-Csv -Path $outputFilePath -Append -Delimiter "," -NoTypeInformation 
            }

        }

    }
}

$endTime = Get-Date 

Write-Host "Output file generated: " $outputFilePath -ForegroundColor Green
Write-Host " Started at" $startTime" and ended at" $endTime 
Write-Host " --- Done ---  "







