'A função ftaSQL é resposavel por intermediar a maior parte das operações com o banco de dados do sistema,
'o tempo de inatividade da sua última utilização define quando o sistema será desconectado'

'---------------------------------------------------------------------------------------------'
'Funcao: ftaSQL                                                                               '
'Por: Valdenir Avila Ramos                                                                    '
'Data: 18/05/2007                                                                             '
'Propósito: Realizar as principais operações do sistema com o banco de dados (FUNÇÃO CRUCIAL) '
'---------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------'                                                                    '
'Modificada na data: 26/07/2022                                                               '
'Por: Thiago Ianzer                                                                           '
'Modificação: Adicionada tratativa para desconectar os usuários por inatividade               '
'---------------------------------------------------------------------------------------------'
Public Function ftaSQL(strSQL As String) As ADODB.Recordset
'    On Error GoTo Erro

    Dim rds As ADODB.Recordset
    Dim isExecSQL As Boolean

    Set rds = CreateObject("ADODB.RecordSet")

    rds.CursorLocation = adUseClient

    isExecSQL = True

'Verifica se tem alguma operação com o banco de dados rodando, se tempo atual > tempo da ultima operação então altera o status para falso
    If Time > modVariaveis.strTempoUltimaOperacao Then
        isExecSQL = False
    End If

    rds.Open strSQL, cnnConexaoSQL, adOpenStatic, adLockOptimistic

    Set ftaSQL = rds


    If isExecSQL = False Then

    'Passa o tempo da consulta anterior, será comparado com o tempo da consulta atual (alimentada no final dessa função)
        Call ftaDesconectaInatividade(modVariaveis.strTempoUltimaOperacao)

    End If

'Captura o tempo do último uso da função (GLOBAL)
    modVariaveis.strTempoUltimaOperacao = Time

    Exit Function
Erro:
    Call ftaTrataErro

End Function


'---------------------------------------------------------------------------------------------'
'**********************************Funções de Inatividade*************************************'
'---------------------------------------------------------------------------------------------'
'Função: Essa série de funções dão suporte a ftaDesconectaInatividade                         '
'Autor: Thiago Ianzer                                                                         '
'Data: 27/10/2022                                                                             '
'Propósito: Permitir funcionalidade de desconexão por inatividade do sistema                  '
'                                                                                             '
'---------------------------------------------------------------------------------------------'

'---------------------------------------------------------------------------------------------'
'Função: ftaDesconectaInatividade                                                             '
'Autor: Thiago Ianzer                                                                         '
'Data: 22/07/2022                                                                             '
'Propósito: Fechar o sistema por tempo de inatividade                                         '
'Observação: É utilizada dentro da função ftaSQL() e ftaSQLRO, recebendo o tempo da ultima    '
'operação da função na variável global modVariaveis.strTempoUltimaOperacao                    '
'---------------------------------------------------------------------------------------------'
Public Function ftaDesconectaInatividade(ByVal strTempo As Date)

If Not blnDebug = True Then On Error GoTo Erro
'Exit Function

'Variáveis de tratativa do tempo
    Dim strTempoUltimoSQL As Date
    Dim strTempoAtual As Date
    Dim strTempoLimite As String
    Dim strTempoDesconexao As Date
    
    strTempoUltimoSQL = strTempo
    strTempoAtual = Time
    'strTempoLimite = "00:00:05"
    strTempoLimite = ftaBuscaParametroInatividade("INATIVIDADE_TEMPO_MAXIMO", "")
    strTempoDesconexao = strTempoAtual - strTempoUltimoSQL
    
    
'Se o tempo da última operação com banco de dados e o tempo limite do parametro forem maior que 00:00:00 e se o usuário não tiver a permissão 01-597 realiza a tratativa
    If strTempoUltimoSQL <> CDate("00:00:00") And ftaVerificaPermissao("01-597") = False Then
         
    'Verifica o tempo para desconexão do sistema, se passar do tempo estabelecido no pelo parametro e fecha o sistema
        If strTempoDesconexao > Format(strTempoLimite, "h:m:s") Then
             
            MsgBox "O tempo máximo de inatividade foi excedido, você será desconectado", vbInformation
             
        'Fecha o módulo desktop e depois encerra o módulo atual
             'Call ftaMatarProcesso("Taura Desktop.exe")
            
            End
             
        End If
       
    End If


    Exit Function
Erro:
   Call ftaTrataErro
   
End Function

'------------------------------------------------------------------'
'Autor: Thiago Ianzer                                              '
'Data: 27/10/2022                                                  '
'Propósito: Buscar valor do parametro (Tempo de Inatividade)       '
'------------------------------------------------------------------'
Public Function ftaBuscaParametroInatividade(strParametroInatividade As String, strFilialTrabalho As String) As String

'Buscando parametro
    Dim rdsParametro As ADODB.Recordset
    
    If strFilialTrabalho = "" Then
        Set rdsParametro = ftaSqlInatividade("Select " & _
                                                "[x002_nome] " & _
                                                ",[x002_valor] " & _
                                             "From " & _
                                                "[x002] (NOLOCK) " & _
                                             "Where " & _
                                                "x002_nome = '" & strParametroInatividade & "'")
    Else
        Set rdsParametro = ftaSqlInatividade("Select " & _
                                                "[x002_nome] " & _
                                                ",[x002_valor] " & _
                                             "From " & _
                                                "[x002] (NOLOCK) " & _
                                             "Where " & _
                                                "x002_nome = '" & strParametroInatividade & "' " & _
                                                "and x002_filial = '" & Trim(strFilTrab) & "'")
    End If
    
    If rdsParametro.RecordCount = 0 Then
        MsgBox "Parametro: " & UCase(strParametroInatividade) & ", não encontrado.", vbInformation
        rdsParametro.Close
    Else
        ftaBuscaParametroInatividade = rdsParametro.Fields(1).Value
        rdsParametro.Close
    End If
        
        
    Exit Function
ftaParametro_Error:
    Call ftaTrataErro

End Function

'------------------------------------------------------------------'
'Autor: Thiago Ianzer                                              '
'Data: 27/10/2022                                                  '
'Propósito: Recebe a consulta que busca o parametro no banco       '
'------------------------------------------------------------------'
Public Function ftaSqlInatividade(strSQL As String) As ADODB.Recordset

'On Error GoTo Erro

    Dim rds As ADODB.Recordset

    Set rds = CreateObject("ADODB.RecordSet")

    rds.CursorLocation = adUseClient

    'rds.Open strSql, cnnConexaoSQL, adOpenStatic, adLockReadOnly
    rds.Open strSQL, cnnConexaoSQL, adOpenForwardOnly, adLockReadOnly


    Set ftaSqlInatividade = rds

    Exit Function
Erro:
    'Call ftaTrataErro
    MsgBox "Erro: " & Err.Number & " - " & Err.Description & vbNewLine & vbNewLine & strSQL, vbCritical

End Function
'---------------------------------------------------------------------------------------------'
'**********************************Funções de Inatividade*************************************'
'---------------------------------------------------------------------------------------------'

