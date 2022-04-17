  strMontaSql = "SELECT " & _
                        "CONVERT(NVARCHAR, a044_data, 103)           AS [Data]  " & _
                        ", a014_cod_und_sec                          AS [Unidade] " & _
                        ", a044_qtd_us                               AS [Quantidade] " & _
                        ", a014_custo_medio_uni                      AS [Custo Medio Uni] " & _
                        ", (a045_saldo_us * a014_custo_medio_uni)    AS [Valor Total] " & _
                        ", (a040_qtd_pri * a014_ultimo_preco_compra) AS [Total Ultimo Preco] " & _
                        ", (a040_qtd_pri * a014_custo_medio_uni)     AS [Total Custo Medio] " & _
                        ", a052_tipo_do_documento                    AS [Tipo de Operacao] " & _
                        ", a052_nro_nf                               AS [Nro Nota] "

        'from
        strMontaSql = strMontaSql & "FROM " & _
                        "a045_estoque_saldos " & _
                        "JOIN a014_produtos              ON (a014_codigo = a045_cod_produto) " & _
                        "JOIN a025_unidades_de_medida    ON (a025_codigo = a014_cod_und_sec) " & _
                        "JOIN a040_pedido_compras_itens  ON (a040_cod_pedido_compra = a045_cod_produto) " & _
                        "JOIN a044_estoque_movimentos    ON (a044_codigo = a045_cod_produto) " & _
                        "JOIN a052_nota_fiscal           ON (a052_codigo = a045_cod_produto) "
                        
        'where
        strMontaSql = strMontaSql & "WHERE " & _
                        "a044_data BETWEEN '" & Format(txtCampo(0).Text, "YYYYMMDD") & "' AND '" & Format(txtCampo(1).Text, "YYYYMMDD") & "'"


--seleciona(filtro)
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