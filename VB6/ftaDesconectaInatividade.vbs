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


'----------------------------------------------------------------------------------------'
'Função: ftaDesconectaInatividade                                                        '
'Autor: Thiago Ianzer                                                                    '
'Data: 22/07/2022                                                                        '
'Propósito: Fechar o sistema por tempo de inatividade                                    '
'Observação: É utilizada dentro da função ftaSQL(), recebendo o tempo da ultima operação '
'da função na variável global modVariaveis.strTempoUltimaOperacao                        '
'----------------------------------------------------------------------------------------'
Public Function ftaDesconectaInatividade(ByVal strTempo As Date)

If Not blnDebug = True Then On Error GoTo Erro
'Variáveis de tratativa do tempo
    Dim strTempoUltimoSQL As Date
    Dim strTempoAtual As Date
    Dim strTempoLimite As String
    Dim strTempoDesconexao As Date
    
    strTempoUltimoSQL = strTempo
    strTempoAtual = Time
    'strTempoLimite = "00:00:30"
    strTempoLimite = ftaParametro("INATIVIDADE_TEMPO_MAXIMO", "")
    strTempoDesconexao = strTempoAtual - strTempoUltimoSQL
    
    
'Se o tempo da última operação com banco de dados e o tempo limite do parametro forem maior que 00:00:00 e o usuário não tiver a permissão 01-597 realiza a tratativa
    If strTempoUltimoSQL <> CDate("00:00:00") And ftaVerificaPermissao("01-597") = False Then
         
    'Verifica o tempo para desconexão do sistema, se passar do tempo estabelecido no pelo parametro fecha o sistema
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