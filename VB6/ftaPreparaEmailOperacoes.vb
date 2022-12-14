
'---------------------------------------------------------------------------------------'
'Função: ftaPreparaEmailOperacoes                                                       '
'Autor: Thiago Ianzer                                                                   '
'Data: 23/09/2022                                                                       '
'Propósito: Prepara os dados do email de acordo com a operação selecionada no sistema   '
'---------------------------------------------------------------------------------------'
Public Sub ftaPreparaEmailOperacoes(strOperacaoEmail As String _
                                   , Optional strCaminhoAnexo As String _
                                   , Optional strCodigoPedido As String)

'O parametro 'strOperacaoEmail' é cadastrado através de um x1 no cadastro de operações de envio de e-mail
If Not blnDebug = True Then On Error GoTo Erro

    Dim rdsGenerico As ADODB.Recordset

'Configurações do email
    Dim strCodigoOperacao As String
    Dim strOperacao As String
    Dim strClienteSMTP As String
    Dim strPortaSMTP As String
    Dim strSSL As String
    Dim strUsuario_email As String
    Dim strSenha_email As String

'Dados mensagem email
    Dim strDe As String
    Dim strPara As String
    Dim strAssunto As String
    Dim strMensagem As String
    Dim strCC As String
    Dim strCCO As String
    Dim strAnexo As String

    'ftaSelecionaCampo


    strConsulta = "SELECT " & _
                        "z019_codigo                 AS [Codigo],   " & _
                        "z019_operacao               AS [Operacao], " & _
                        "z018_smtp_cliente           AS [Cliente],  " & _
                        "z018_smtp_porta             AS [Porta],    " & _
                        "z018_smtp_ssl               AS [SSL],      " & _
                        "z018_email                  AS [Email De], " & _
                        "z018_senha                  AS [Senha],    " & _
                        "z019_para                   AS [Para],     " & _
                        "ISNULL(z019_assunto, '')    AS [Assunto],  " & _
                        "ISNULL(z019_mensagem, '')   AS [Mensagem], " & _
                        "CASE " & _
                            "WHEN z019_CC = NULL THEN ISNULL(z019_CC, '')          " & _
                            "WHEN z019_CC = '-'  THEN REPLACE('-', z019_CC, '' )   " & _
                            "ELSE z019_CC " & _
                        "END     AS [CC], " & _
                        "CASE " & _
                            "WHEN z019_CCO = NULL THEN ISNULL(z019_CCO, '')        " & _
                            "WHEN z019_CCO = '-'  THEN REPLACE('-', z019_CCO, '' ) " & _
                            "ELSE z019_CCO " & _
                        "END     AS [CCO] "

    strConsulta = strConsulta & _
                    "FROM " & _
                        "z019_operacoes_email        (NOLOCK) " & _
                        "LEFT JOIN z018_contas_email (NOLOCK) ON (z019_cod_origem = z018_codigo) " & _
                    "WHERE " & _
                        "z019_operacao = '" & UCase(strOperacaoEmail) & "' "


    Set rdsGenerico = ftaSQL(CStr(strConsulta))


'Verifica se a operação existe no banco de dados
    If rdsGenerico.EOF Then
        If modVariaveis.strNomUser = "NEY" Then
            MsgBox "Operação de e-mail não encontrada no sistema", vbExclamation
        End If
        Exit Sub
    End If


'Dados de servidor e conta
    strCodigoOperacao = rdsGenerico.Fields("Codigo").Value   'Codigo operação
    strOperacao = rdsGenerico.Fields("Operacao").Value       'Operação para a qual o email está cadastrado
    strClienteSMTP = rdsGenerico.Fields("Cliente").Value     'Servidor SMTP Cliente
    strPortaSMTP = rdsGenerico.Fields("Porta").Value         'Servidor SMTP Porta
    strSSL = rdsGenerico.Fields("SSL").Value                 'Servidor SMTP SSL
    strUsuario_email = rdsGenerico.Fields("Email De").Value  'Email de Envio: Usuário
    strSenha_email = rdsGenerico.Fields("Senha").Value       'Email de Envio: Senha

'Dados de envio
    'strDe = "<" & rdsGenerico.Fields("Email De").Value & ">"
    strDe = rdsGenerico.Fields("Email De").Value                                    'Email Remetente
    strPara = rdsGenerico.Fields("Para").Value                                      'Email Destinatário
    strAssunto = rdsGenerico.Fields("Assunto").Value & " - (NÃO RESPONDA)"          'Email Assunto
    strMensagem = rdsGenerico.Fields("Mensagem").Value & " - (NÃO RESPONDA)"        'Email Mensagem
    strCC = rdsGenerico.Fields("CC").Value                                          'Email CC (com cópia)
    strCCO = rdsGenerico.Fields("CCO").Value                                        'Email CCO (com cópia oculta)
    strAnexo = ""                                                                   'Email Anexo (arquivos em anexo)


'Descriptografa a senha do email
    strSenha_email = ftaDeCriptSenha(strSenha_email)

'Realiza tratativa com base na operação de email selecionada
    Select Case UCase(strOperacao)

        Case "ALTERAÇÃO STATUS PEDIDO COMPRA"
            
        'Utiliza o código do pedido para pegar o email do solicitante e os dados do recebindo dos itens do pedido
            strConsulta = "SELECT " & _
                                " a039_codigo                                         AS [codigo],             " & _
                                " ISNULL(a151_codigo, 'NENHUM')                       AS [sc],                 " & _
                                " ISNULL(a010_email, '')                              AS [email],              " & _
                                " a014_codigo                                         AS [cod_produto],        " & _
                                " ISNULL(a014_descricao, '(NENHUM)')                  AS [descricao],          " & _
                                " ISNULL(a040_qtd_pri, '0')                           AS [qtd_pedido],         " & _
                                " ISNULL(a154_qtd_pri, '0')                           AS [qtd_recebida],       " & _
                                " ISNULL(a040_qtd_pri - a154_qtd_pri, a040_qtd_pri)   AS [qtd_falta_receber],  " & _
                                " a153_usuario                                        AS [usuario_recebeu],    " & _
                                " a153_data_hora                                      AS [data_recebimento]    " & _
                            "FROM " & _
                                "a040_pedido_compras_itens                      (NOLOCK) " & _
                                "LEFT JOIN a014_produtos                        (NOLOCK) ON (a014_codigo = a040_cod_produto)            " & _
                                "LEFT JOIN a039_pedido_compras                  (NOLOCK) ON (a040_cod_pedido_compra = a039_codigo)      " & _
                                "LEFT JOIN a010_pessoas                         (NOLOCK) ON (a010_codigo = a039_cod_solicitante)        " & _
                                "LEFT JOIN a153_recebimento_pedido_compra_dv    (NOLOCK) ON (a039_codigo = a153_cod_pedido_compra)      " & _
                                "LEFT JOIN a154_recebimento_itens               (NOLOCK) ON (a040_codigo = a154_cod_item_pedido_compra) " & _
                                "LEFT JOIN a151_solicitacao_compra              (NOLOCK) ON (a151_codigo = a039_cod_doc_org)            " & _
                            "WHERE " & _
                                "a039_codigo = '" & strCodigoPedido & "' "
        
            
            Set rdsGenerico = ftaSQL(CStr(strConsulta))
            

        'Verifica se o pedido existe
            If Not rdsGenerico.EOF Then
            
            'Verifica se o solicitante possui um email cadastrado antes de continuar a operação
                If rdsGenerico.Fields("email").Value = "-" Or rdsGenerico.Fields("email").Value = "" Then
                    Exit Sub
                Else
                    strPara = rdsGenerico.Fields("email").Value
                    
                'Removendo caracteres especiais do email e espaços
                    strPara = Trim(ftaRegexEmail(strPara))
                    strPara = Replace(strPara, ";", ",")
                    
                'Verifica se o email é valido
                    If ftaValidaEmail(strPara) = False Then
                        If modVariaveis.strNomUser = "NEY" Then
                            MsgBox "Endereço de e-mail do solicitante é inválido!", vbExclamation
                        End If
                        Exit Sub
                    End If
                    
                'Inserindo tabela HTML pro envio de email
                    strMensagem = strMensagem & _
                                      "<!DOCTYPE HTML> " & _
                                      "<html> " & _
                                      "<head> " & _
                                        "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""> " & _
                                      "</head> " & _
                                      "<body> " & _
                                        "<br>" & _
                                        "<br>" & _
                                        "<h3>SOLICITAÇÃO DE COMPRA(" & rdsGenerico.Fields("sc").Value & ")</h3>" & _
                                        "<br> " & _
                                        "<h3>ITENS DO PEDIDO(" & rdsGenerico.Fields("codigo").Value & "): </h3>" & _
                                        "<br>" & _
                                        "<table border=1>" & _
                                        "<tr>" & _
                                          "<th>" & "ID Produto &nbsp;</th>" & _
                                          "<th>" & "Descrição &nbsp;</th>" & _
                                          "<th>" & "Qtd. Pedido &nbsp;</th>" & _
                                          "<th>" & "Qtd. Recebida &nbsp;</th>" & _
                                          "<th>" & "Qtd. A Receber &nbsp;</th>" & _
                                          "<th>" & "Usuário &nbsp;</th>" & _
                                          "<th>" & "Data &nbsp;</th>" & _
                                        "</tr>"


                    Do While Not rdsGenerico.EOF
                        strMensagem = strMensagem & _
                                        "<tr>" & _
                                          "<td>" & rdsGenerico.Fields("cod_produto").Value & "&nbsp;</td>" & _
                                          "<td>" & rdsGenerico.Fields("descricao").Value & "&nbsp;</td>" & _
                                          "<td>" & rdsGenerico.Fields("qtd_pedido").Value & "&nbsp;</td>" & _
                                          "<td>" & rdsGenerico.Fields("qtd_recebida").Value & "&nbsp;</td>" & _
                                          "<td>" & rdsGenerico.Fields("qtd_falta_receber").Value & "&nbsp;</td>" & _
                                          "<td>" & rdsGenerico.Fields("usuario_recebeu").Value & "&nbsp;</td>" & _
                                          "<td>" & rdsGenerico.Fields("data_recebimento").Value & "&nbsp;</td>" & _
                                        "</tr>"
                                      
                        rdsGenerico.MoveNext
                    Loop
                End If
            Else
                Exit Sub
            End If
            
            strMensagem = strMensagem & _
                                        "</table> " & _
                                    "</body> " & _
                                    "</html>"
                                    
            
        'Arquivo de anexo para testes
            'strAnexo = "C:\Users\Estação 3\Pictures\article3.jpg"
            
            'Enviando e-mail para o solicitante
            Call modFuncoes.ftaEnviarEmailCDO(strClienteSMTP, _
                                              CInt(strPortaSMTP), _
                                              CInt(strSSL), _
                                              strUsuario_email, _
                                              strSenha_email, _
                                              strPara, _
                                              strDe, _
                                              strAssunto, _
                                              strMensagem, _
                                              strCC, _
                                              strCCO, _
                                              strAnexo)
       
        Case Else
        
        'Enviando e-mail
            Call modFuncoes.ftaEnviarEmailCDO(strClienteSMTP, _
                                              CInt(strPortaSMTP), _
                                              CInt(strSSL), _
                                              strUsuario_email, _
                                              strSenha_email, _
                                              strPara, _
                                              strDe, _
                                              strAssunto, _
                                              strMensagem, _
                                              strCC, _
                                              strCCO, _
                                              strAnexo)
                                              
    End Select


    Exit Sub
Erro:
    Call ftaTrataErro

End Sub


'---------------------------------------------------------------------------------------'
'*******************************Funções de Criptografia*********************************'
'---------------------------------------------------------------------------------------'
'Função: Essa série de funções tem o objetivo de criptografar e descriptografar strings '
'Autor: Thiago Ianzer                                                                   '
'Data: 21/07/2022                                                                       '
'Propósito: Criptografar strings, principalmente senhas dos usuários                    '
'                                                                                       '
'---------------------------------------------------------------------------------------'

'---------------------------------------------------------------------------------------'
'Função: ftaCriptSenha                                                                  '
'Autor: Thiago Ianzer                                                                   '
'Propósito: Criptografar uma string                                                     '
'---------------------------------------------------------------------------------------'
Public Function ftaCriptSenha(Psenha As String) As Variant
If Not blnDebug = True Then On Error GoTo Erro
    
    Dim v_sqlerrm As String
    Dim SenhaCript As String
    Dim var1 As String
    
    Const MIN_ASC = 32
    Const MAX_ASC = 126
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    
    chave = 2001 ''qualquer nº para montar o algorítimo da criptografia
    Dim offset As Long
    Dim str_len As Integer
    Dim i As Integer
    Dim ch As Integer
        
    to_text = ""
    offset = ftaNumericPassword(chave)
    Rnd -1
    Randomize offset
    str_len = Len(Psenha)
    For i = 1 To str_len
        ch = Asc(Mid$(Psenha, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch + offset) Mod NUM_ASC)
            ch = ch + MIN_ASC
            to_text = to_text & Chr$(ch)
        End If
    Next i
    
    ftaCriptSenha = to_text
    
    
    Exit Function
Erro:
    Call ftaTrataErro
End Function

'---------------------------------------------------------------------------------------'
'Função: ftaDeCriptSenha                                                                '
'Autor: Thiago Ianzer                                                                   '
'Propósito: Descriptografar a string gerada por ftaCriptSenha                           '
'---------------------------------------------------------------------------------------'
Public Function ftaDeCriptSenha(Psenha As String) As Variant
If Not blnDebug = True Then On Error GoTo Erro

    Dim v_sqlerrm As String
    Dim SenhaCript As String
    Dim var1 As String
    
    Const MIN_ASC = 32  ' Space.
    Const MAX_ASC = 126 ' ~.
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    
    chave = 2001 ''qualquer nº para montar o algorítimo da criptografia
    Dim offset As Long
    Dim str_len As Integer
    Dim i As Integer
    Dim ch As Integer
     
    to_text = ""
    offset = ftaNumericPassword(chave)
    Rnd -1
    Randomize offset
    str_len = Len(Psenha)
    For i = 1 To str_len
        ch = Asc(Mid$(Psenha, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            to_text = to_text & Chr$(ch)
        End If
    Next i
    
    ftaDeCriptSenha = to_text
    

    Exit Function
Erro:
    Call ftaTrataErro
End Function

'---------------------------------------------------------------------------------------'
'Função: ftaNumericPassword                                                             '
'Autor: Thiago Ianzer                                                                   '
'Propósito: Auxiliar na Criptografia e Descriptografia de strings                       '
'---------------------------------------------------------------------------------------'
Private Function ftaNumericPassword(ByVal password As String) As Long
If Not blnDebug = True Then On Error GoTo Erro
    
    Dim Value As Long
    Dim ch As Long
    Dim shift1 As Long
    Dim shift2 As Long
    Dim i As Integer
    Dim str_len As Integer

    str_len = Len(password)
    For i = 1 To str_len
        ' Adiciona a próxima letra
        ch = Asc(Mid$(password, i, 1))
        Value = Value Xor (ch * 2 ^ shift1)
        Value = Value Xor (ch * 2 ^ shift2)

        ' Change the shift offsets.
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    
    ftaNumericPassword = Value
    
    
    Exit Function
Erro:
    Call ftaTrataErro
End Function
'-------------------------------------------------------------------------------------------'
'*******************************FIM Funções de Criptografia*********************************'
'-------------------------------------------------------------------------------------------'

