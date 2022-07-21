
'---------------------------------------------------------------------------------------'
'Autor     : Thiago                                                                     '
'Data      : 15/06/2022                                                                 '
'Propósito : Realizar validação da abertura e fechamento de estoque                     '
'---------------------------------------------------------------------------------------'
Dim isValidaAberturaEstoque(4) As Boolean
		
isValidaAberturaEstoque(0) = ftaValidaDatas()
isValidaAberturaEstoque(1) = ftaValidaDataAberta()
isValidaAberturaEstoque(2) = ftaValidaSemEstoqueAberto()
isValidaAberturaEstoque(3) = ftaValidaDataMaiorUltimoEstoque()


For i = 0 To 3
	If isValidaAberturaEstoque(i) = "False" Then
		Exit Sub
	End If
Next


'------------------------------------------------------------------------------------------'
'Procedures : Validação da abertura e fechamento do estoque                                '
' Autor     : Thiago                                                                       '
' Data      : 15/06/2022                                                                   '
' Propósito : Condições para validar a abertura de estoque                                 '
' Detalhes  : Foram criadas funções para validar a abertura e o fechamento do estoque      '
'------------------------------------------------------------------------------------------'
Public Function ftaValidaDatas() As String
If Not blnDebug = True Then On Error GoTo Erro

    ftaValidaDatas = False

    'Dim isValidaEstoque As Boolean
    Dim rdsVerificaData As ADODB.Recordset
        
     'Montando condições
    Set rdsVerificaData = ftaSQLRO("select " & _
                                        "a046_estado " & _
                                        ", a046_data " & _
                                        ", a046_usuario " & _
                                        ", a046_datatime " & _
                                    "From " & _
                                        "a046_estoque  (NOLOCK) " & _
                                    "Where " & _
                                        "a046_data >= '" & Format(txtCampo(2).Text, "YYYYMMDD") & "' " & _
                                        "and a046_filial = '" & txtCampo(3).Text & "'")
                                        
         
    If Not rdsVerificaData.EOF Then

        'Valida abertura de estoque com Datas Anteriores ou Igual a Selecionada
        MsgBox "O estoque de " & Format(rdsVerificaData.Fields("a046_data").Value, "DD/MM/YYYY") & _
                " foi " & UCase(rdsVerificaData.Fields("a046_estado").Value) & "." & vbNewLine & vbNewLine & _
                "   Por: " & rdsVerificaData.Fields("a046_usuario").Value & vbNewLine & vbNewLine & _
                "   Em: " & rdsVerificaData.Fields("a046_datatime").Value & vbNewLine & vbNewLine & _
                "   IMPOSSÍVEL REALIZAR OPERAÇÃO!", vbCritical
        
        ftaValidaDatas = False
        
    Else
        ftaValidaDatas = True
    End If
                                                                 
        ftaValidaDatas = ftaValidaDatas
                                                                 
    Exit Function
Erro:
    Call ftaTrataErro
End Function

Public Function ftaValidaDataAberta() As String
If Not blnDebug = True Then On Error GoTo Erro

    'valor padrão da função
    ftaValidaDataAberta = False
    
    'Dim isValidaEstoque As Boolean
    Dim rdsVerificaData As ADODB.Recordset
     
     
    'Valida Abertura de estoque com alguma Data em com estado Aberto
    Set rdsVerificaData = ftaSQLRO("select " & _
                                        "a046_estado " & _
                                        ", a046_data " & _
                                        ", a046_usuario " & _
                                        ", a046_datatime " & _
                                    "From " & _
                                        "a046_estoque  (NOLOCK) " & _
                                    "Where " & _
                                        "a046_estado = 'ABERTO' " & _
                                        "and a046_filial = '" & txtCampo(3).Text & "'")
        
    If Not rdsVerificaData.EOF Then
    
        'Valida abertura de estoque com Datas Anteriores ou Igual a Selecionada
        
        MsgBox "O estoque de " & Format(rdsVerificaData.Fields("a046_data").Value, "DD/MM/YYYY") & _
                " foi " & UCase(rdsVerificaData.Fields("a046_estado").Value) & "." & vbNewLine & vbNewLine & _
                "   Por: " & rdsVerificaData.Fields("a046_usuario").Value & vbNewLine & vbNewLine & _
                "   Em: " & rdsVerificaData.Fields("a046_datatime").Value & vbNewLine & vbNewLine & _
                "   IMPOSSÍVEL REALIZAR OPERAÇÃO!", vbCritical
        
        ftaValidaDataAberta = False
        
    Else
        ftaValidaDataAberta = True
    End If
                                            
        ftaValidaDataAberta = ftaValidaDataAberta
                                                            
    Exit Function
Erro:
    Call ftaTrataErro
End Function

Public Function ftaValidaSemEstoqueAberto() As String
If Not blnDebug = True Then On Error GoTo Erro

    'valor padrão da função
    ftaValidaSemEstoqueAberto = False

    'Dim isValidaEstoque As Boolean
    Dim rdsVerificaData As ADODB.Recordset

     'Valida Abertura de estoque sem estoque aberto no dia anterior.
    If ftaParametro("VALIDA_FECHAMENTO_ESTOQUE_DIARIAMENTE", "") = "SIM" Then
         Set rdsVerificaData = ftaSQLRO(" select " & _
                                            " a046_estado " & _
                                            " , a046_data " & _
                                            " , a046_usuario " & _
                                            " , a046_datatime " & _
                                        " From " & _
                                            " a046_estoque (NOLOCK) " & _
                                        " Where " & _
                                            " a046_data = DATEADD(day, -1, '" & Format(txtCampo(2).Text, "YYYYMMDD") & "' ) " & _
                                            " and a046_filial = '01' ")
                                            
        If rdsVerificaData.EOF Then
        
            'Valida abertura de estoque com Datas Anteriores ou Igual a Selecionada
            
            MsgBox "O estoque do dia anterior à " & txtCampo(2).Text & " não foi encontrado. " & vbNewLine & vbNewLine & _
                    "   IMPOSSÍVEL REALIZAR OPERAÇÃO!", vbCritical
            
            ftaValidaSemEstoqueAberto = False
        End If
        
    Else
    
        ftaValidaSemEstoqueAberto = True
    End If
    
    ftaValidaSemEstoqueAberto = ftaValidaSemEstoqueAberto
                                                            
    Exit Function
Erro:
    Call ftaTrataErro
End Function

Public Function ftaValidaDataMaiorUltimoEstoque() As Boolean
If Not blnDebug = True Then On Error GoTo Erro

    'valor padrão da função
    ftaValidaDataMaiorUltimoEstoque = False

    'Dim isValidaEstoque As Boolean
    Dim rdsVerificaData As ADODB.Recordset
    
    'Adicinado por: Thiago | Data: 30/05/2022
    Set rdsVerificaData = ftaSQLRO("select " & _
                                        "TOP 1 " & _
                                        "CONVERT(varchar, a046_data, 103) AS [Última data] " & _
                                    "FROM " & _
                                        "a046_estoque  (NOLOCK) " & _
                                    "ORDER BY " & _
                                        "a046_data DESC ")

    Dim strDataUltimoEstoque As String
    Dim isValidaDataEstoque As Boolean
    
    strDataUltimoEstoque = rdsVerificaData.Fields("Última data")
    isValidaDataEstoque = CDate(txtCampo(2).Text) > CDate(strDataUltimoEstoque) + 1
    
    'Valida Abertura de estoque, data do estoque não pode ser maior que a data posterior do último estoque
    If isValidaDataEstoque = True Then
        MsgBox "A data para abertura de estoque é inválida " & vbNewLine & vbNewLine & _
                        "   IMPOSSÍVEL REALIZAR OPERAÇÃO!", vbCritical
        
        ftaValidaDataMaiorUltimoEstoque = False
    Else
    
        ftaValidaDataMaiorUltimoEstoque = True
    End If
    
        ftaValidaDataMaiorUltimoEstoque = ftaValidaDataMaiorUltimoEstoque
                                                            
    Exit Function
Erro:
    Call ftaTrataErro
End Function
