-- Capture the results and save it to CopyIntoTablesConfig.csv 

Declare @OutputDirPrefix varchar(50)= 'C:\migratemaster\output\GenerageCopyIntoScripts' 
Declare @CopyType varchar(20) = 'BlobMiCsv' -- Blob Managed Identity CSV Type 
--Declare @CopyType varchar(20) = 'DfsMiCsv' --  Data Lake, Managed Identity CSV Type 

Select '1' as Active,
db_name()  as DatabaseName, 
s.name as SchemaName, 
t.name as TableName, 
s.name + '_asa' as AsaSchema,
@OutputDirPrefix + '\' + db_name() + '\' + @CopyType as SqlFilePath
from sys.tables t 
inner join sys.schemas s 
on t.schema_id = s.schema_id 
inner join sys.databases d
on d.name = db_name()  and t.type_desc = 'USER_TABLE' 
and t.temporal_type_desc ='NON_TEMPORAL_TABLE' 
and t.object_id not in (select object_id from sys.external_tables)