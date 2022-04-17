'Uma das querys criadas para a geração de relatórios no sistema
   'Corpo
        strMontaSql = "SELECT " & _
                        "a014_codigo                                 As [Id Item] " & _
                        ", CONVERT(NVARCHAR, a044_data, 103)         AS [Data]  " & _
                        ", a014_descricao                            AS [Descricao] " & _
                        ", a014_end_almox                            AS [Local] " & _
                        ", a014_custo_medio_uni                      AS [Custo Medio Uni] " & _
                        ", a045_saldo_us                             AS [Saldo Us]  " & _
                        ", (a040_qtd_pri * a014_ultimo_preco_compra) AS [Total Ultimo Preco]  " & _
                        ", (a040_qtd_pri * a014_custo_medio_uni)     AS [Total Custo Medio] " & _
                        ", (a045_saldo_us * a014_custo_medio_uni)    AS [Valor Total] "

        'from
        strMontaSql = strMontaSql & "FROM " & _
                                        "a045_estoque_saldos " & _
                                        "JOIN a014_produtos              ON (a014_codigo = a045_cod_produto) " & _
                                        "JOIN a040_pedido_compras_itens     ON (a040_cod_pedido_compra = a045_cod_produto) " & _
                                        "JOIN a044_estoque_movimentos    ON (a044_codigo = a045_cod_produto) "
                        
        'where
        strMontaSql = strMontaSql & "WHERE " & _
                                        "a044_data BETWEEN '" & Format(txtCampo(0).Text, "YYYYMMDD") & "' AND '" & Format(txtCampo(1).Text, "YYYYMMDD") & "'"

--select(filtro)
 Case 3
         Call ftaSelecaoDeDados("SELECT " & _
                                    "a014_codigo            As [ID Prod] " & _
                                    ", a014_descricao       AS [Descricao] " & _
                                "From " & _
                                    "a045_estoque_saldos " & _
                                    "JOIN a014_produtos ON (a014_codigo = a045_cod_produto) " & _
                                "Group By " & _
                                    "a014_codigo " & _
                                    ", a014_descricao " & _
                                "Order By " & _
                                    "a014_descricao ", txtCampo(2), txtDescricao(2))
