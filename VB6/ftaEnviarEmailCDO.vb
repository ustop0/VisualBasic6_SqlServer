'---------------------------------------------------------------------------------------'
'Função: ftaEnviarEmailCDO                                                              '
'Autor: Thiago Ianzer                                                                   '
'Data: 04/07/2022                                                                       '
'Propósito: Enviar com ou sem anexo                                                     '
'Observação: É necessário o componente Microsoft CDO to Windows 2000 Library            '
'---------------------------------------------------------------------------------------'
Public Function ftaEnviarEmailCDO(strClienteSMTP As String _
                                 , strPortaSMTP As Integer _
                                 , strSSL As Integer _
                                 , strUsuario_email As String _
                                 , strSenha_email As String _
                                 , strRemetente_email As String _
                                 , strNomeRemetente_email As String _
                                 , strAssunto_email As String _
                                 , strMenssagem_email As String _
                                 , Optional ByVal strCC_email As String _
                                 , Optional ByVal strCCO_email As String _
                                 , Optional ByVal strAnexo_email As String)
    
If Not blnDebug = True Then On Error GoTo Erro

'Removendo caracteres especiais do email e espaços
    strUsuario_email = Trim(ftaRegexEmail(strUsuario_email))
    strUsuario_email = Replace(strUsuario_email, ";", ",")
    
    strRemetente_email = Trim(ftaRegexEmail(strRemetente_email))
    strRemetente_email = Replace(strRemetente_email, ";", ",")

    strNomeRemetente_email = Trim(ftaRegexEmail(strNomeRemetente_email))
    strNomeRemetente_email = Replace(strNomeRemetente_email, ";", ",")
    
    
    On Error Resume Next 'Set up error checking
    Set cdoMsg = CreateObject("CDO.Message")
    Set cdoConf = CreateObject("CDO.Configuration")
    Set cdoFields = cdoConf.Fields
    'Send one copy with Google SMTP server (with autentication)
    schema = "http://schemas.microsoft.com/cdo/configuration/"
    
'Configurações do envio de email
    cdoFields.Item(schema & "sendusing") = 2
    cdoFields.Item(schema & "smtpserver") = strClienteSMTP
    cdoFields.Item(schema & "smtpserverport") = strPortaSMTP
    cdoFields.Item(schema & "smtpauthenticate") = 1
    cdoFields.Item(schema & "sendusername") = strUsuario_email
    cdoFields.Item(schema & "sendpassword") = strSenha_email
    cdoFields.Item(schema & "smtpusessl") = strSSL
    cdoFields.Update
    
    
'Verifica se foi passado anexo
    With cdoMsg
         .To = strRemetente_email
         .From = "<" & strNomeRemetente_email & ">"
         .Subject = strAssunto_email
         .HTMLBody = strMenssagem_email ' Body of message can be any HTML code
            
    'Verificando copia oculta e anexo
         If strCC_email <> "" Then .CC = strCC_email
         If strCCO_email <> "" Then .BCC = strCCO_email
         If strAnexo_email <> "" Then .AddAttachment strAnexo_email
            
         Set .Configuration = cdoConf
         .send
    End With
  
   
'Verifica envio
'    If Err.Number = 0 Then
'          MsgBox "Email enviado com sucesso", , "Email"
'    Else
'          MsgBox "Erro ao enviar o email" & Err.Number, , "Email"
'    End If
    
    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set cdoFields = Nothing
    
    Exit Function
Erro:
    Call ftaTrataErro
End Function
                      
