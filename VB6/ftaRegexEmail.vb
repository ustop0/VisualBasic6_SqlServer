'Essa função foi desenvolvida para evitar problemas com characteres ocultos, muitos usuários copiam os dados
'de email para cadastrar o sistema e comumente acabam vindo esses characteres, limpando para o cadastro.

'regex campo de email
Private Sub Command1_Click()
    Dim result As String
    Dim strEmail As String

    strEmail = Text1(0).Text
    
    result = ftaRegexEmail(strEmail)

    Text1(1).Text = result
End Sub

'---------------------------------------------------------------------------------------'
' Procedure : ftaRegexEmail                                                             '
' Autor     : Thiago                                                                    '
' Data      : 14/03/2022                                                                '
' Propósito : Remove caracteres especiais email (regex)                                 '
'                                                                                       '
'---------------------------------------------------------------------------------------'

Private Function ftaRegexEmail(strEmail As String) As String
If Not blnDebug = True Then On Error GoTo Erro

    Dim strPattern As String
    Dim regex As New RegExp

    strPattern = "[`~!#$%^&*()_|+\-=?:'¨ \n\t\b,<>\{\}\[\]\\\/]"

    regex.Global = True
    regex.Pattern = strPattern


    result = regex.Replace(strEmail, "")

    ftaRegexEmail = result

    Exit Function
Erro:
    Call ftaTrataErro
End Function
