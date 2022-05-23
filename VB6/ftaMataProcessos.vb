'Essa função serve para matar um determinado processo em execução, veio da necessidade de fechar
'o módulo de faturamento do sistema, pois quando é substituido o por uma nova versão no servidor do cliente
'é encessário que todas as instâncias sejam encerradas para não ocorrer problemas

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As String) As Long
Private Declare Function WMI Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As String) As Long

'---------------------------------------------------------------
'Função matar processos

' Variaveis para usar Wmi
Dim ListaProcessos As Object
Dim ObjetoWMI As Object
Dim ProcessoAEncerrar As Object
  
Private Function ftaMatarProcesso( _
    strNomeProcesso As String, _
    Optional strSim As Boolean = True) As Boolean
    
    If Not blnDebug = True Then On Error GoTo Erro
    
    Dim strRecebeProcesso As String
  
    Set ObjetoWMI = GetObject("winmgmts:")
  
    If IsNull(ObjetoWMI) = False Then
        ' Nesta variavel obtemos os processos
        Set ListaProcessos = ObjetoWMI.InstancesOf("win32_process")
        
        For Each ProcessoAEncerrar In ListaProcessos
            If ProcessoAEncerrar.Name = strNomeProcesso Then
                strRecebeProcesso = ProcessoAEncerrar.Name
                Text2.Text = strRecebeProcesso
            End If
        Next
        
        If strRecebeProcesso = "" Then
            MsgBox "O processo não está em execução"
            Debug.Print "O processo não está em execução"
            Exit Function
        End If
    End If
  
    'Eliminamos as variaveis objeto
    Set ListaProcessos = Nothing
    Set ObjetoWMI = Nothing
  
    ftaMatarProcesso = False
  
    Set ObjetoWMI = GetObject("winmgmts:")
  
    If IsNull(ObjetoWMI) = False Then
  
    'instanciamos a variavel
    Set ListaProcessos = ObjetoWMI.InstancesOf("win32_process")
  
    For Each ProcessoAEncerrar In ListaProcessos
        If UCase(ProcessoAEncerrar.Name) = UCase(strNomeProcesso) Then
            ProcessoAEncerrar.Terminate (0)
            ftaMatarProcesso = True
        End If
  
    Next
    End If
  
    'Elimina as variaveis
    Set ListaProcessos = Nothing
    Set ObjetoWMI = Nothing
    
Exit Function
Erro:
    Call ftaTrataErro
    Debug.Print "Erro ao encerrar processo"
    
End Function

'---------------------------------------------------------------
'btn matar processos API Windows
Private Sub btnMataProcesso_Click()
    
    'Eliminamos as variaveis objeto
    Set ListaProcesos = Nothing
    Set ObjetoWMI = Nothing

    Dim strNomeProcesso As String
    
    strNomeProcesso = "notepad.exe"
    
    ftaMatarProcesso LCase$(strNomeProcesso), True

End Sub
