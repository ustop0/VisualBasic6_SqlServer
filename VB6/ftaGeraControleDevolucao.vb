

'-----------------------------------------------------------------------------------------------------------'
'Função: ftaGeraControleDevolucao                                                                           '
'Autor: Thiago Ianzer                                                                                       '
'Data: 18/11/2022                                                                                           '
'Propósito: Gerar devoluções através de operações do sistema (Função generalizada para diversas operações)  '
'   1 - Gerar na pesagem de ordem de recebimento: ORDEM_RECEBIMENTO_PESAGEM                                 '
'   2 -                                                                                                     '
'-----------------------------------------------------------------------------------------------------------'
Public Function ftaGeraControleDevolucao(strOperacao As String, _
                                         strCodigoPedido As String)

If Not blnDebug = True Then On Error GoTo Erro
  
    Dim rdsBuscaItens As ADODB.Recordset
    Dim rdsItensApontamento As ADODB.Recordset
    Dim strBuscaItens As String
    Dim strInsereItens As String
    
'Gera ID e Código
    Dim strCodigo As String
    Dim strSequencial As String
'Itens Insert
    Dim strData As String
    Dim strCodPedido As String
    Dim strMotivo As String
    Dim strCodProduto As String
    Dim strQtdPri As String
    Dim strQtdSec As String
    Dim strValorUnitario As String
    Dim strCodLote As String
    Dim strObs As String
    Dim strDatahora As String
    Dim strUsuario As String
    Dim strGeraMovimentacao As String
        
        
        
    strBuscaItens = "SELECT " & _
                       "b033_data_hora                                  AS [data],             " & _
                       "a067_cod_doc_origem                             AS [cod_pedido],       " & _
                       "ISNULL(a068_motivo, '-')                        AS [motivo],           " & _
                       "a068_cod_produto                                AS [cod_produto],      " & _
                       "b033_qtd_pri                                    AS [qtd_pri],          " & _
                       "b033_qtd_sec                                    AS [qtd_sec],          " & _
                       "a068_valor_unitario                             AS [valor_unitario],   " & _
                       "b033_cod_lote                                   AS [cod_lote],         " & _
                       "a067_observacoes                                AS [observacoes],      " & _
                       "b033_data_hora                                  AS [data_hora],        " & _
                       "b033_usuario                                    AS [usuario],          " & _
                       "a068_gera_movimentacao                          AS [gera_movimentacao] " & _
                   "FROM " & _
                       "b033_apontamentos_ord_rec            (NOLOCK) " & _
                       "JOIN a067_ordem_de_recebimento       (NOLOCK) ON (b033_cod_ordem_recebimento = a067_codigo) " & _
                       "JOIN a068_ordem_de_recebimento_itens (NOLOCK) ON (a068_cod_ordem_recebimento = a067_codigo) " & _
                   "WHERE " & _
                       "a068_cod_ordem_recebimento = '" & strCodigoPedido & "' "
    
    
    Set rdsBuscaItens = ftaSQL(CStr(strBuscaItens))
    
    
    
'Inserindo os itens na b016_devoluções
    If Not rdsBuscaItens.EOF Then
    
        Select Case strOperacao
        
            Case "ORDEM_RECEBIMENTO_PESAGEM"
                
                Do While Not rdsBuscaItens.EOF
                          
                    'Gera ID e Código
                    strCodigo = ftaGeraCodigo("b016_devolucoes")
                    strSequencial = CDbl(Mid(strCodigo, 4))
                    'Outros dados
                    strData = rdsBuscaItens.Fields("data").Value
                    strCodPedido = rdsBuscaItens.Fields("cod_pedido").Value
                    strMotivo = rdsBuscaItens.Fields("motivo").Value
                    strCodProduto = rdsBuscaItens.Fields("cod_produto").Value
                    strQtdPri = rdsBuscaItens.Fields("qtd_pri").Value
                    strQtdSec = rdsBuscaItens.Fields("qtd_sec").Value
                    strValorUnitario = rdsBuscaItens.Fields("valor_unitario").Value
                    strCodLote = rdsBuscaItens.Fields("cod_lote").Value
                    strObs = rdsBuscaItens.Fields("observacoes").Value
                    strDatahora = rdsBuscaItens.Fields("data_hora").Value
                    strUsuario = rdsBuscaItens.Fields("usuario").Value
                    strGeraMovimentacao = rdsBuscaItens.Fields("gera_movimentacao").Value
                                  
                          
                    'Colunas
                     strInsereItens = "INSERT INTO b016_devolucoes" & _
                                                  "( " & _
                                                    "b016_codigo,           " & _
                                                    "b016_sequencial,       " & _
                                                    "b016_data,             " & _
                                                    "b016_cod_pedido,       " & _
                                                    "b016_motivo,           " & _
                                                    "b016_cod_produto,      " & _
                                                    "b016_qtd_pri,          " & _
                                                    "b016_qtd_sec,          " & _
                                                    "b016_valor_unitario,   " & _
                                                    "b016_cod_lote,         " & _
                                                    "b016_obs,              " & _
                                                    "b016_data_hora,        " & _
                                                    "b016_usuario,          " & _
                                                    "b016_gera_movimentacao " & _
                                                  " ) "
                                        
                     'Valores
                     strInsereItens = strInsereItens & _
                                        "VALUES " & _
                                                "( " & _
                                                  " '" & strCodigo & "',                                " & _
                                                  " '" & strSequencial & "',                            " & _
                                                  " '" & Format(strData, "yyyymmdd") & "',              " & _
                                                  " '" & strCodPedido & "',                             " & _
                                                  " '" & strMotivo & "',                                " & _
                                                  " '" & strCodProduto & "',                            " & _
                                                  " '" & strQtdPri & "',                                " & _
                                                  " '" & strQtdSec & "',                                " & _
                                                  " '" & strValorUnitario & "',                         " & _
                                                  " '" & strCodLote & "',                               " & _
                                                  " '" & strObs & "',                                   " & _
                                                  " '" & Format(strDatahora, "yyyymmdd hh:mm:ss") & "', " & _
                                                  " '" & strUsuario & "',                               " & _
                                                  " '" & strGeraMovimentacao & "'                       " & _
                                                " ) "
                                                  
        
                    Set rdsInsereItens = ftaSQL(CStr(strInsereItens))
                
                    rdsBuscaItens.MoveNext
                Loop
                
                
            'Finalizando a Ordem de recebimento
                Call ftaSQL("UPDATE a067_ordem_de_recebimento SET a067_estado = 'FN' WHERE a067_codigo = '" & strCodigoPedido & "' ")
        
        End Select
    
    End If
    

    Exit Function
Erro:
   Call ftaTrataErro
End Function


