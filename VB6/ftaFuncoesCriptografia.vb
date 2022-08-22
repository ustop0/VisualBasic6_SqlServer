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