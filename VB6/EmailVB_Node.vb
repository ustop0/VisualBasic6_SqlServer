'Essas função foram criadas para a integrar o sistema a uma api desenvolvida em node, porém ocorreram problemas
'na conversão de arquivos Danfe em PDF para binário de base 64, foram testadas diversas variações de funções
', tanto por banco de dados SQL SERVER quanto pelo VB, nenhuma funcionou corretamente, tendo como um dos problemas exceder
'o limite de characteres da string do VB, o resto da função funciona perfeitamente

'função base64
'Private Function ftaEncodeBase64(ByRef arrData() As Byte) As String 'converte para base64

    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement

    Set objXML = New MSXML2.DOMDocument

'    ' byte array para base64
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    ftaEncodeBase64 = objNode.Text

    Set objNode = Nothing
    Set objXML = Nothing
End Function

'funcao enviar email node
Public Function ftaEnviaEmailNode(strUsuario As String _
                                    , strSenha As String _
                                    , strFrom As String _
                                    , strEmail As String _
                                    , strAssunto As String _
                                    , strMensagem As String _
                                    , Optional strAnexo As String) As String

If Not blnDebug = True Then On Error GoTo Erro

    Dim http As Object
    Dim strURL As String
    Dim strDadosEmail As String
    Dim rdsBase64 As ADODB.Recordset


    Dim b() As Byte
    Dim filenum As Long
    filenum = FreeFile()

    Open "C:\testePDF.pdf" For Binary Access Read As filenum
    ReDim b(1 To LOF(filenum))
    Get filenum, 1, b()
    Close filenum

    'Converte base64 pelo banco
    Set rdsBase64 = ftaSQLRO("Select " & _
                                     "dbo.fncBase64_Encode(' & b & ') AS [base64]")

    strAnexo = rdsBase64.Fields("base64").Value

'    strAnexo = ftaEncodeBase64(b)
'    strAnexo = "teste email"

    strDadosEmail = strUsuario & "/" & strSenha & "/" & strFrom & "/" & strEmail & "/" & strAssunto & "/" & strMensagem & "/" & strAnexo


    'URL do modulo email da api node
    strURL = "http://192.168.1.105:2300/enviarEmail/" & strDadosEmail

    Set http = CreateObject("WinHttp.WinHttprequest.5.1")

    With http
              .Open "POST", strURL, False

              .setRequestHeader "Content-Type", "text/plain;charset=UTF-8"
              .send strDadosEmail

              strResponse = .responseText

              MsgBox strResponse

              ftaEnviaEmailNode = strResponse

     End With

    Exit Function
Erro:
    Call ftaTrataErro

End Function
