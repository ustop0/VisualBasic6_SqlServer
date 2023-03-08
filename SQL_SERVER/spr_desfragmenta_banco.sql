
			CREATE PROCEDURE [dbo].[spr_desfragmenta_banco]
			AS
			BEGIN

				/*
				'---------------------------------------------------------------------------------------
				' Autor     : Thiago IAnzer
				' Data      : 16/12/2022
				' Propósito : Desfragmentar o banco de dados (reorganiza  e reconstrói os indices das tabelas)
				'---------------------------------------------------------------------------------------
				' Atualizações:
					| Data		| Por					| Descrição																											|
					------------+-----------------------+--------------------------------------------------------------------------------------------------------------------
					16/12/2022	  Thiago Ianzer			  A procedure foi revisada e melhorada para aumentar a sua eficiência
				*/

				DECLARE @nome_tabela NVARCHAR(300)
				DECLARE @nome_indice NVARCHAR(500)
				DECLARE @fragmentacao_media NVARCHAR(300)
				DECLARE @script NVARCHAR(600)

				DECLARE @cont INTEGER
				DECLARE @cont_not INTEGER

				SET @cont = 0
				SET @cont_not = 0

				-- Cursor para percorrer os registros
				DECLARE desfragmenta_banco CURSOR FOR


					-- Consultar a fragmentação média:
					SELECT
						( object_name(b.object_id) )		AS [nome_tabela],
						ISNULL( name, '(NENHUM INDICE)')	AS [nome_indice],
						avg_fragmentation_in_percent		AS [fragmentacao_media],
						ISNULL( (CASE
									WHEN avg_fragmentation_in_percent > 30 THEN 'ALTER index ' + name + ' ON ' + object_name(b.object_id) + ' REBUILD'
									WHEN avg_fragmentation_in_percent >= 5 and avg_fragmentation_in_percent <= 30 THEN 'ALTER index ' + name + ' ON ' + object_name(b.object_id) + ' REORGANIZE'
								END), '(NENHUM SCRIPT)')	AS [script]
					FROM 
						sys.dm_db_index_physical_stats (db_id('curso'), null, null, null, null) AS a -- (Parâmetros da função: banco de dados, tabela, indice, partição física, modo de analise: default, null, limited (limitado), sampled (amostra), detailed (detalhado))
						JOIN sys.indexes AS b ON (a.object_id = b.object_id and a.index_id = b.index_id)
					WHERE
						avg_fragmentation_in_percent <> 0	AND
						NAME <> '(NENHUM INDICE)'			AND
						(CASE
							WHEN avg_fragmentation_in_percent > 30 THEN 'ALTER index ' + name + ' ON ' + object_name(b.object_id) + ' REBUILD'
							WHEN avg_fragmentation_in_percent >= 5 and avg_fragmentation_in_percent <= 30 THEN 'ALTER index ' + name + ' ON ' + object_name(b.object_id) + ' REORGANIZE'
						 END) <> '(NENHUM SCRIPT)'
					ORDER BY
						avg_fragmentation_in_percent DESC

		
				--Abrindo Cursor
				OPEN desfragmenta_banco
 
				-- Lendo a próxima linha
				FETCH NEXT FROM desfragmenta_banco INTO @nome_tabela, @nome_indice, @fragmentacao_media, @script
 
				-- Percorrendo linhas do cursor (enquanto houverem)
				WHILE @@FETCH_STATUS = 0
				BEGIN
 
				-- Executando as rotinas de reorganização e reconstrução dos indices
					BEGIN TRY

						EXEC( @script )
						If @fragmentacao_media > 30
							BEGIN
								PRINT 'Fragmentação: ' + @fragmentacao_media
								PRINT 'Executado: ' + @script 
								PRINT 'Indice: ' + @nome_indice + ' da tabela ' + @nome_tabela + ' Reconstruído com sucesso'
								PRINT '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
								PRINT '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
								PRINT ''
							END
						ELSE IF @fragmentacao_media >= 5 AND @fragmentacao_media <= 30
							BEGIN
								PRINT 'Fragmentação: ' + @fragmentacao_media
								PRINT 'Executado: ' + @script 
								PRINT 'Indice: ' + @nome_indice + ' da tabela ' + @nome_tabela + ' Reorganizado com sucesso'
								PRINT '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
								PRINT '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
								PRINT ''
							END
						
						SET @cont = @cont + 1

					END TRY
					BEGIN CATCH
						PRINT 'Fragmentação: ' + @fragmentacao_media
						PRINT 'Não é possível executar uma operação para o índice ' + @nome_indice + ' da tabela ' + @nome_tabela + ' porque o índice contém a colunas de tipo de dados text, ntext, image ou FILESTREAM.'
						PRINT '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
						PRINT '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
						PRINT ''
						SET @cont_not = @cont_not + 1

					END CATCH;

	
				-- Lendo a próxima linha
				FETCH NEXT FROM desfragmenta_banco INTO @nome_tabela, @nome_indice, @fragmentacao_media, @script
				END
 
				-- Fechando Cursor para leitura
				CLOSE desfragmenta_banco
 
				-- Finalizado o cursor
				DEALLOCATE desfragmenta_banco

				PRINT ''
				PRINT 'Indices Alterados: '						+ CAST( @cont AS NVARCHAR(20) )
				PRINT 'Indices que não puderam ser Alterados: ' + CAST( @cont_not AS NVARCHAR(20) )
				PRINT 'Indices Total: '							+ CAST( ( @cont + @cont_not) AS NVARCHAR(20) )

			END
			GO
	
