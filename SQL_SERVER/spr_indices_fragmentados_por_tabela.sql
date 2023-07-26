USE <NOME DO BANCO>
GO

CREATE PROCEDURE [dbo].[spr_indices_fragmentados_por_tabela]
	@tabela_filtro NVARCHAR(300) = NULL  -- Parâmetro opcional para filtrar as tabelas
AS
BEGIN
	/*
		'---------------------------------------------------------------------------------------
		' Autor     : Thiago Ianzer
		' Data      : 26/07/2023
		' Propósito : Exibe todos os índices por tabela e cria os scripts para "REBUILD" ou "REORGANIZE" para todos os índices por tabela do banco de dados
		'---------------------------------------------------------------------------------------
		' Atualizações:
			| Data		| Por					| Descrição																											|
			------------+-----------------------+--------------------------------------------------------------------------------------------------------------------
	*/

	SELECT
		(OBJECT_NAME(b.object_id))      AS [nome_tabela],
		ISNULL(name, '(NENHUM INDICE)') AS [nome_indice],
		avg_fragmentation_in_percent    AS [fragmentacao_media],
		(CASE
			WHEN avg_fragmentation_in_percent > 30 THEN 'ALTER INDEX ' + QUOTENAME(name) + ' ON ' + QUOTENAME(OBJECT_NAME(b.object_id)) + ' REBUILD'
			WHEN avg_fragmentation_in_percent >= 5 AND avg_fragmentation_in_percent <= 30 THEN 'ALTER INDEX ' + QUOTENAME(name) + ' ON ' + QUOTENAME(OBJECT_NAME(b.object_id)) + ' REORGANIZE'
			END)                            AS [script]
	FROM 
		sys.dm_db_index_physical_stats (db_id(), NULL, NULL, NULL, NULL) AS a
		JOIN sys.indexes AS b ON (a.object_id = b.object_id AND a.index_id = b.index_id)
	WHERE 
		avg_fragmentation_in_percent <> 0
		AND name <> '(NENHUM INDICE)'
		AND (@tabela_filtro IS NULL OR OBJECT_NAME(b.object_id) = @tabela_filtro)  -- Aplica o filtro de tabela quando o parâmetro não for nulo

END
GO


