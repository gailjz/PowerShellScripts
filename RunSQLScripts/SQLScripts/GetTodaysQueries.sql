--Alter Database AdventureWorksDW2017
--SET QUERY_STORE = ON (OPERATION_MODE = READ_WRITE);
SELECT Txt.query_text_id
,Txt.query_sql_text 
,Pl.plan_id
,Qry.query_parameterization_type_desc
,Qry.last_execution_time
FROM sys.query_store_plan AS Pl
INNER JOIN sys.query_store_query AS Qry
    ON Pl.query_id = Qry.query_id
INNER JOIN sys.query_store_query_text AS Txt
    ON Qry.query_text_id = Txt.query_text_id 
where Qry.last_execution_time >= DATEADD(DAY, -1, getdate())
Order By Qry.last_execution_time   DESC -- ASC 

--Select * from sys.query_store_query 
