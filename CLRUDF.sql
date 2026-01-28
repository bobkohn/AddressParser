USE ISCHTR
GO

drop FUNCTION dbo.clrSQLParseAll
go

drop ASSEMBLY AddressParser
go

CREATE ASSEMBLY AddressParser
FROM '\\std-sql01\SQLAssemblies\AddressParser.DLL'
go

CREATE FUNCTION dbo.clrSQLParseAll
(	@indata nvarchar(4000)
)	RETURNS nvarchar(4000)
WITH EXECUTE AS CALLER 
AS 
EXTERNAL NAME AddressParser.[AddressParser.Parser].SQLParseAll
GO

