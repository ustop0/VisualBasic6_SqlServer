/*
				'---------------------------------------------------------------------------------------
				' Autor     : Thiago Ianzer
				' Data      : 14/12/2022
				' Propósito : Mover os dados da tabela z014_log para outro banco de dados auxiliar, após isso limpar a tabela do banco do cliente num período abaixo de 90 dias
				'---------------------------------------------------------------------------------------
				' Atualizações:
					| Data		| Por					| Descrição																											|
					------------+-----------------------+--------------------------------------------------------------------------------------------------------------------

			*/

			CREATE PROCEDURE [dbo].[spr_backup_limpa_z014_log]
			AS   
			BEGIN 
	
				INSERT INTO dbTauraSGF_Backup_Log.dbo.z014_log  SELECT * FROM z014_log WHERE z014_data_hora <= ( GETDATE() - 90 ) ORDER BY z014_data_hora DESC

	
				DECLARE @data_log NVARCHAR(20)
				DECLARE @retorno_1 NVARCHAR(20)


				SET @data_log = ( SELECT TOP 1 z014_data_hora FROM z014_log WHERE z014_data_hora <= ( GETDATE() - 90 ) ORDER BY z014_data_hora DESC )

				IF @data_log <= ( GETDATE() - 90 )
	            
					SET @retorno_1 = 'TRUE'

				ELSE
					SET @retorno_1 = 'FALSO'


				IF @retorno_1 = 'TRUE'

					DELETE FROM 
						z014_log 
					WHERE 
						z014_data_hora <= ( GETDATE() - 90 ) 
  
			END
			GO
