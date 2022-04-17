/** Essa função foi criada para converter arquivos pdf de Danfes para binário de base64, não funciona corretamente **/

-- criptografando BASE 64
CREATE FUNCTION [dbo].[ftaBase64_codifica] (
@string VARCHAR(MAX)
) 
RETURNS VARCHAR(MAX)
AS BEGIN

	DECLARE 
		@source VARBINARY(MAX), 
		@encoded VARCHAR(MAX)

	SET @source = CONVERT(VARBINARY(MAX), @string)
	SET @encoded = CAST('' AS XML).value('xs:base64Binary(sql:variable("@source"))', 'varchar(max)')
	RETURN @encoded
END

--executa função
SELECT dbo.ftaBase64_codifica('teste')


--importando pdf
Select 'C:\testePDF.pdf', BulkColumn 
FROM Openrowset( Bulk 'C:\testePDF.pdf', Single_Blob) as Arquivo 
