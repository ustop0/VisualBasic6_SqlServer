

'--------------------------------------------------------------------------------------'
'Autor: Thiago Ianzer                                                                  '
'Data: 10/11/2022                                                                      '
'Propósito: Buscar os itens do pedido de compra diversos e insere nos itens da O.S,    '
'(1)Insert                                                                             '
'--------------------------------------------------------------------------------------'
Public Function ftaItensPedCompraDiversos()

    Dim rdsBuscaItens As ADODB.Recordset
    Dim strBuscaItens As String
    Dim strInsereItens As String
    Dim strCodOrdemServico As String
    Dim strCodPedido As String
    
    
    strCodOrdemServico = txtCampo(0).Text
    strCodPedido = txtCampo(11).Text
    
'Busca os itens do pedido de compra diversos
    strBuscaItens = "SELECT " & _
                         "A014_CODIGO         AS [ID],               " & _
                         "A014_SEQUENCIAL     AS [CODIGO],           " & _
                         "A014_COD_BARRAS     AS [COD. INTEGRACAO],  " & _
                         "A014_DESCRICAO      AS [DESCRICAO],        " & _
                         "A014_PRECO_COMPRA   AS [PREÇO COMPRA],     " & _
                         "A040_COD_TES        AS [TES],              " & _
                         "A040_QTD_PRI        AS [QTD PRI],          " & _
                         "A040_QTD_SEC        AS [QTD SEC],          " & _
                         "A040_PRECO_UNI      AS [PRECO UNI],        " & _
                         "A040_DESCONTO       AS [DESCONTO],         " & _
                         "A040_TOTAL_ITEM     AS [TOTAL ITEM],       " & _
                         "A040_OBSERVACOES    AS [OBSERVACOES]       " & _
                    "FROM " & _
                         "A014_PRODUTOS                   (NOLOCK) " & _
                         "JOIN A026_FAMILIAS_PRODUTOS     (NOLOCK) ON (A026_CODIGO = A014_COD_FAMILIA)  " & _
                         "JOIN a040_pedido_compras_itens  (NOLOCK) ON (a040_cod_produto = a014_codigo)  " & _
                    "WHERE " & _
                         "A040_COD_PEDIDO_COMPRA = '" & strCodPedido & "' " & _
                         "AND A014_ESTADO_CADASTRO = 'ATIVO'              " & _
                         "AND A036_PERMITE_PEDIDO_COMPRA_DIVERSOS = 'SIM' "
                        
                        
    Set rdsBuscaItens = ftaSQL(CStr(strBuscaItens))
    
    
'Insert/Update itens na ordem de serviço
    If Not rdsBuscaItens.EOF Then
        
        
        Dim strCodProduto As String
        Dim strCodTes As String
        Dim strCodContaGerencial As String
        Dim strQtdPri As String
        Dim strQtdSec As String
        Dim strPrecoUni As String 'Talvez seja criado um parametro para selecionar qual o preço uni
        Dim strPrecoDesconto As String
        Dim strTotalItem As String
        Dim strObs As String
            
            
    'Insert/Update itens na ordem de serviço
        Do While Not rdsBuscaItens.EOF
        
            strCodItem = ""
            strSequencial = ""
            
            'Gerando o codigo dos itens da O.S e buscando os itens do pedido de compra diversos
            If strCodItem = "" Then
                strCodItem = ftaGeraCodigo("a050_ordem_servicos_itens")
                strSequencial = CDbl(Mid(strCodItem, 4))
            End If
    
            strCodProduto = rdsBuscaItens.Fields("ID").Value
            strCodTes = rdsBuscaItens.Fields("TES").Value
            strCodContaGerencial = "00-0"
            strQtdPri = rdsBuscaItens.Fields("QTD PRI").Value
            strQtdSec = rdsBuscaItens.Fields("QTD SEC").Value
            strPrecoUni = rdsBuscaItens.Fields("PRECO UNI").Value
            strPrecoDesconto = rdsBuscaItens.Fields("DESCONTO").Value
            strTotalItem = rdsBuscaItens.Fields("TOTAL ITEM").Value
            strObs = rdsBuscaItens.Fields("OBSERVACOES").Value
                          
                          
            strInsereItens = "INSERT INTO a050_ordem_servicos_itens  " & _
                                        "(a050_codigo,               " & _
                                        " a050_sequencial,           " & _
                                        " a050_cod_ordem_de_servico, " & _
                                        " a050_cod_produto,          " & _
                                        " a050_cod_tes,              " & _
                                        " a050_cod_conta_gerencial,  " & _
                                        " a050_qtd_pri,              " & _
                                        " a050_qtd_sec,              " & _
                                        " a050_preco_uni,            " & _
                                        " a050_desconto,             " & _
                                        " a050_total_item,           " & _
                                        " a050_obs                   " & _
                                        " ) "
                                        
            strInsereItens = strInsereItens & _
                            "VALUES " & _
                                    "( '" & strCodItem & "',           " & _
                                    "  '" & strSequencial & "',        " & _
                                    "  '" & strCodOrdemServico & "',   " & _
                                    "  '" & strCodProduto & "',        " & _
                                    "  '" & strCodTes & "',            " & _
                                    "  '" & strCodContaGerencial & "', " & _
                                    "  '" & strQtdPri & "',            " & _
                                    "  '" & strQtdSec & "',            " & _
                                    "  '" & strPrecoUni & "',          " & _
                                    "  '" & strPrecoDesconto & "',     " & _
                                    "  '" & strTotalItem & "',         " & _
                                    "  '" & strObs & "'                " & _
                                    " ) "
    
            
            Call ftaSQL(CStr(strInsereItens))
                          
            rdsBuscaItens.MoveNext
        Loop
    
    End If

End Function


'*********** UTILIZANDO A FUNÇÃO - ftaItensPedCompraDiversos() ***********'

'--------------------------------------------------------------------------------------'
'Autor: Thiago Ianzer                                                                  '
'Data: 10/11/2022                                                                      '
'Propósito: Buscar os itens do pedido de compra diversos e insererir nos itens da O.S  '
'--------------------------------------------------------------------------------------'
    Case 10
          
    'Verifica se foi selecionada uma placa
        If txtCampo(13).Text = txtCampo(13).Tag Then
            MsgBox "Nenhuma placa selecionada!", vbInformation
            Exit Sub
        End If
        
    'Verifica se foi selecionado um responsável
        If txtCampo(13).Text = txtCampo(13).Tag Then
            MsgBox "Nenhum responsável selecionado!", vbInformation
            Exit Sub
        End If
        
        
    'Se a ordem de serviço já existir
        Dim rdsBuscaItens As ADODB.Recordset
        Dim strBuscaItens As String
        
    'Seleciona o pedido de compra diversos (c/ centros de custo de frota)
        strBuscaItens = "SELECT " & _
                            "a039_codigo                                AS [ID PEDIDO],         " & _
                            "a039_usuario                               AS [USUÁRIO],           " & _
                            "a010_codigo                                AS [ID FORNECEDOR],     " & _
                            "a010_nome                                  AS [FORNECEDOR],        " & _
                            "ISNULL(a030_codigo, '00-0')                AS [ID],                " & _
                            "ISNULL(a030_numero, '00.000.000')          AS [NÚMERO],            " & _
                            "ISNULL(a030_nome, '(NENHUM)')              AS [CENTRO DE CUSTO],   " & _
                            "ISNULL( SUM(a040_total_item), '0.00' )     AS [VALOR PEDIDO],      " & _
                            "a039_estado                                AS [ESTADO],            " & _
                            "a039_nro_nf_entrada                        AS [NRO NF ENTRADA],    " & _
                            "a039_modalidade_do_frete                   AS [MOD. FRETE],        " & _
                            "a039_data_hora                             AS [DATA],              " & _
                            "a039_data_previsao_entrega                 AS [DATA PREV. ENTREGA] "
                        
        strBuscaItens = strBuscaItens & _
                        "FROM " & _
                            "a039_pedido_compras                 (NOLOCK) " & _
                            "LEFT JOIN a010_pessoas              (NOLOCK) ON (a010_codigo = a039_cod_fornecedor)    " & _
                            "LEFT JOIN a040_pedido_compras_itens (NOLOCK) ON (a040_cod_pedido_compra = a039_codigo) " & _
                            "LEFT JOIN a030_centro_custo         (NOLOCK) ON (a030_codigo = a039_cod_centro_custo)  " & _
                        "WHERE " & _
                            "a039_estado <> 'FT'         " & _
                        "GROUP BY " & _
                            "a039_codigo,                " & _
                            "a039_sequencial,            " & _
                            "a039_usuario,               " & _
                            "a010_codigo,                " & _
                            "a010_nome,                  " & _
                            "a039_nro_nf_entrada,        " & _
                            "a039_estado,                " & _
                            "a039_modalidade_do_frete,   " & _
                            "a039_data_hora,             " & _
                            "a039_data_previsao_entrega, " & _
                            "a030_codigo,                " & _
                            "a030_numero,                " & _
                           "a030_nome                   " & _
                        "ORDER BY " & _
                            "a039_sequencial DESC "
                            
                            '"AND a030_numero IN ('03.003',  '04.003', '05.001') " & _

         
         Set rdsBuscaItens = ftaSelecaoDeDados2(CStr(strBuscaItens))
        
    If Not rdsBuscaItens.EOF Then
            
            'Incluir ou editar os itens da O.S
            If txtCampo(0).Text <> "" Then
            
                'Incluindo O.S no sistema e listando os itens na grid
                If txtCampo(7).Text <> "FN" Then
                   
                    'Quando o pedido de compra diversos é alterado os itens antigos são deletados da ordem de serviço
                    If txtCampo(11).Text <> txtCampo(11).Tag Then
                        
                        'Campos com os dados do pedido
                        txtCampo(11).Text = rdsBuscaItens.Fields("ID PEDIDO").Value
                        txtDescricao(5).Text = rdsBuscaItens.Fields("USUÁRIO").Value
                        
                        'Campos com os dados do fornecedor
                        txtCampo(2).Text = rdsBuscaItens.Fields("ID FORNECEDOR").Value
                        txtDescricao(0).Text = rdsBuscaItens.Fields("FORNECEDOR").Value
                        
                        'Campos com os dados do Centro de Custo
                        txtCampo(15).Text = rdsBuscaItens.Fields("ID").Value
                        txtDescricao(2).Text = rdsBuscaItens.Fields("CENTRO DE CUSTO").Value
                        
                        Call ftaSQL("DELETE FROM a050_ordem_servicos_itens WHERE a050_cod_ordem_de_servico = '" & txtCampo(0).Text & "' ")
                        Call ftaItensPedCompraDiversos
                        Call ftaBuscaItens
                        
                    End If
                
                End If
                
            ElseIf txtCampo(0).Text = "" Then
                                    
                'Se o código não tiver sido gerado
                txtCampo(0).Text = ftaGeraCodigo(strTabela)
                txtCampo(1).Text = CDbl(Mid(txtCampo(0).Text, 4))
                
                'Campos com os dados do pedido
                txtCampo(11).Text = rdsBuscaItens.Fields("ID PEDIDO").Value
                txtDescricao(5).Text = rdsBuscaItens.Fields("USUÁRIO").Value
                
                'Campo com os dados do fornecedor
                txtCampo(2).Text = rdsBuscaItens.Fields("ID FORNECEDOR").Value
                txtDescricao(0).Text = rdsBuscaItens.Fields("FORNECEDOR").Value
            
                'Campos com os dados do Centro de Custo
                txtCampo(15).Text = rdsBuscaItens.Fields("ID").Value
                txtDescricao(2).Text = rdsBuscaItens.Fields("CENTRO DE CUSTO").Value
                
                'Inserindo itens na a050
                Call ftaItensPedCompraDiversos
                'Lista os itens da a050
                Call ftaBuscaItens
            
                'Inserir
                intGravarAutomatico = 1
                Call cmdConfirmar_Click
                Call ftaBuscaItens
                
            End If
        
        End If

