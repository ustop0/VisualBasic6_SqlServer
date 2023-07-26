USE <NOME DO BANCO>
GO

CREATE PROCEDURE [dbo].[spr_desfragmenta_banco2] 
    @tabela_filtro NVARCHAR(300) = NULL  -- Par�metro opcional para filtrar as tabelas
AS
BEGIN

	/*
		'---------------------------------------------------------------------------------------
		' Autor     : Thiago Ianzer
		' Data      : 26/07/2023
		' Prop�sito : Desfragmentar o banco de dados (reorganiza  e reconstr�i os indices das tabelas)
		'---------------------------------------------------------------------------------------
		' Atualiza��es:
			| Data		| Por					| Descri��o																											|
			------------+-----------------------+--------------------------------------------------------------------------------------------------------------------
	*/

    -- Cria��o de uma tabela tempor�ria para armazenar os comandos de desfragmenta��o
    DECLARE @IndexCommands TABLE (
        ID INT IDENTITY(1, 1) PRIMARY KEY,
        Command NVARCHAR(MAX)
    )

    -- Cria��o de uma tabela tempor�ria para armazenar os �ndices n�o alterados
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


    -- Armazenar os comandos de desfragmenta��o na vari�vel de tabela
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
        AND (@tabela_filtro IS NULL OR OBJECT_NAME(b.object_id) = @tabela_filtro)  -- Aplica o filtro de tabela quando o par�metro n�o for nulo


    -- Percorrendo a vari�vel de tabela com os comandos de desfragmenta��o
    DECLARE @index INT = 1
    DECLARE @totalCount INT = (SELECT COUNT(*) FROM @IndexCommands)

    WHILE @index <= @totalCount
    BEGIN
        SELECT @script = Command FROM @IndexCommands WHERE ID = @index

        BEGIN TRY
            EXEC (@script)
        END TRY
        BEGIN CATCH
            -- Se ocorrer um erro, insere o �ndice n�o alterado na tabela de �ndices n�o alterados
            INSERT INTO @FailedIndexes (TableName, IndexName, ErrorMessage)
            VALUES (@nome_tabela, @nome_indice, ERROR_MESSAGE())
        END CATCH

        SET @index = @index + 1
    END

    -- Exibe os �ndices n�o alterados e suas mensagens de erro (se houver)
    DECLARE @failedCount INT = (SELECT COUNT(*) FROM @FailedIndexes)

    IF @failedCount > 0
    BEGIN
        PRINT '�ndices n�o alterados devido a erros:'
        SELECT TableName, IndexName, ErrorMessage FROM @FailedIndexes
    END

    -- Exibe as informa��es
    PRINT ' '
    PRINT 'Desfragmenta��o conclu�da.'
    PRINT '�ndices Alterados: ' + CAST((@totalCount - @failedCount) AS NVARCHAR(20))
    PRINT '�ndices N�o Alterados: ' + CAST(@failedCount AS NVARCHAR(20))
    PRINT 'N�mero Total de �ndices: ' + CAST(@totalCount AS NVARCHAR(20))
END
GO


