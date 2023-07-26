USE <NOME DO BANCO>
GO

CREATE PROCEDURE [dbo].[spr_desfragmenta_banco2] 
    @tabela_filtro NVARCHAR(300) = NULL  -- Parâmetro opcional para filtrar as tabelas
AS
BEGIN

	/*
		'---------------------------------------------------------------------------------------
		' Autor     : Thiago Ianzer
		' Data      : 26/07/2023
		' Propósito : Desfragmentar o banco de dados (reorganiza  e reconstrói os indices das tabelas)
		'---------------------------------------------------------------------------------------
		' Atualizações:
			| Data		| Por					| Descrição																											|
			------------+-----------------------+--------------------------------------------------------------------------------------------------------------------
	*/

    -- Criação de uma tabela temporária para armazenar os comandos de desfragmentação
    DECLARE @IndexCommands TABLE (
        ID INT IDENTITY(1, 1) PRIMARY KEY,
        Command NVARCHAR(MAX)
    )

    -- Criação de uma tabela temporária para armazenar os índices não alterados
    DECLARE @FailedIndexes TABLE (
        ID INT IDENTITY(1, 1) PRIMARY KEY,
        TableName NVARCHAR(300),
        IndexName NVARCHAR(500),
        ErrorMessage NVARCHAR(MAX)
    )

    DECLARE @nome_tabela NVARCHAR(300)
    DECLARE @nome_indice NVARCHAR(500)
    DECLARE @fragmentacao_media FLOAT
    DECLARE @script NVARCHAR(MAX)
    DECLARE @command NVARCHAR(MAX)


    -- Armazenar os comandos de desfragmentação na variável de tabela
    INSERT INTO @IndexCommands (Command)
    SELECT
        (CASE
            WHEN avg_fragmentation_in_percent > 30 THEN 'ALTER INDEX ' + QUOTENAME(name) + ' ON ' + QUOTENAME(OBJECT_NAME(b.object_id)) + ' REBUILD'
            WHEN avg_fragmentation_in_percent >= 5 AND avg_fragmentation_in_percent <= 30 THEN 'ALTER INDEX ' + QUOTENAME(name) + ' ON ' + QUOTENAME(OBJECT_NAME(b.object_id)) + ' REORGANIZE'
         END)
    FROM 
        sys.dm_db_index_physical_stats (db_id(), NULL, NULL, NULL, NULL) AS a
        JOIN sys.indexes AS b ON (a.object_id = b.object_id AND a.index_id = b.index_id)
    WHERE 
        avg_fragmentation_in_percent <> 0
        AND name <> '(NENHUM INDICE)'
        AND (@tabela_filtro IS NULL OR OBJECT_NAME(b.object_id) = @tabela_filtro)  -- Aplica o filtro de tabela quando o parâmetro não for nulo


    -- Percorrendo a variável de tabela com os comandos de desfragmentação
    DECLARE @index INT = 1
    DECLARE @totalCount INT = (SELECT COUNT(*) FROM @IndexCommands)

    WHILE @index <= @totalCount
    BEGIN
        SELECT @script = Command FROM @IndexCommands WHERE ID = @index

        BEGIN TRY
            EXEC (@script)
        END TRY
        BEGIN CATCH
            -- Se ocorrer um erro, insere o índice não alterado na tabela de índices não alterados
            INSERT INTO @FailedIndexes (TableName, IndexName, ErrorMessage)
            VALUES (@nome_tabela, @nome_indice, ERROR_MESSAGE())
        END CATCH

        SET @index = @index + 1
    END

    -- Exibe os índices não alterados e suas mensagens de erro (se houver)
    DECLARE @failedCount INT = (SELECT COUNT(*) FROM @FailedIndexes)

    IF @failedCount > 0
    BEGIN
        PRINT 'Índices não alterados devido a erros:'
        SELECT TableName, IndexName, ErrorMessage FROM @FailedIndexes
    END

    -- Exibe as informações
    PRINT ' '
    PRINT 'Desfragmentação concluída.'
    PRINT 'Índices Alterados: ' + CAST((@totalCount - @failedCount) AS NVARCHAR(20))
    PRINT 'Índices Não Alterados: ' + CAST(@failedCount AS NVARCHAR(20))
    PRINT 'Número Total de Índices: ' + CAST(@totalCount AS NVARCHAR(20))
END
GO


