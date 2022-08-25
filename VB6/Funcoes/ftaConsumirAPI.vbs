

'********************************************************************
'************************FUNÇÃO CONSUMIR APIs************************
'********************************************************************
'---------------------------------------------------------------------------------------'
'Função: ftaConsumirAPI                                                                 '
'Autor: Thiago Ianzer                                                                   '
'Data: 05/08/2022                                                                       '
'Propósito: Consumir Web APIs                                                           '
'---------------------------------------------------------------------------------------'
Public Function ftaConsumirAPI(strURL_API As String _
                                , strMetodoEnvio As String _
                                , strBody As String _
                                , Optional strContentType As String _
                                , Optional strAnexo As String) As String

If Not blnDebug = True Then On Error GoTo Erro

    Dim http As Object
    'Dim rdsBase64 As ADODB.Recordset


'Verifica se o tipo de conteudo foi preenchido
    If IsEmpty(strContentType) Then

        strContentType = "text/plain;charset=UTF-8"
    End If


    Set http = CreateObject("WinHttp.WinHttprequest.5.1")

'Enviando os dados e consumindo a API
    With http
              .Open strMetodoEnvio, strURL, False

              .setRequestHeader "Content-Type", strContentType
              .send strBody

              strResponse = .responseText

              MsgBox strResponse

              ftaConsumirAPI = strResponse

     End With

    'ftaConsumirAPI = strResponse

    Exit Function
Erro:
    Call ftaTrataErro

End Function

