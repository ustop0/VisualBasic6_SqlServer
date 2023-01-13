	CREATE PROCEDURE [dbo].[spr_limpa_arquivo_de_log_banco]
			AS
			BEGIN

				DECLARE @tamanho_banco NVARCHAR(100)
				DECLARE @tamanho_banco_log NVARCHAR(100)
				DECLARE @tamanho_reduzir NVARCHAR(100)

				SET @tamanho_banco = (SELECT 
										'size' = convert(nvarchar(15), convert (bigint, size) * 8) + N'' 
										FROM 
											sysfiles 
										WHERE 
											name NOT LIKE '%%_log')

				SET @tamanho_banco_log = (SELECT 
											'size' = convert(nvarchar(15), convert (bigint, size) * 8) + N'' 
											FROM 
												sysfiles 
											WHERE 
												name LIKE '%%_log')


				SET @tamanho_reduzir = CAST( @tamanho_banco AS INTEGER )/2 --1352


				IF CAST( @tamanho_banco_log AS INTEGER ) >= CAST( @tamanho_reduzir AS INTEGER )

					--Pega o log do banco
					DECLARE @banco_nome sysname
					DECLARE @banco_log NVARCHAR(50)
					DECLARE @sql NVARCHAR(MAX)

					SET @banco_nome = ( SELECT db_name() )
					SET @banco_log = ( SELECT name FROM sysfiles WHERE name LIKE '%%_log' )
	
	
					SET @sql = 'USE ' + @banco_nome +
									  + ' ALTER DATABASE ' + @banco_nome + ' SET RECOVERY SIMPLE; ' 
									  + ' DBCC SHRINKFILE (' + @banco_log + ', 1);' 
									  + ' ALTER DATABASE ' + @banco_nome + ' SET RECOVERY FULL; '


					PRINT @sql

					EXEC( @sql );
	
			END			
			GO
