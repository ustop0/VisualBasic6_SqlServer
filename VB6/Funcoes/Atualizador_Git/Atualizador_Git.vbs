
'--------------------'
'***COMANDOS SHELL***'
'--------------------'
Private Sub cmdShell_Click()
    
'Caminho CMD: "C:\Windows\System32\cmd.exe /c/git_teste.sh"
'Caminho script: strCMD = "C:\Fontes\ShellGit.sh"
    
    Dim strCMD As String

'Terminal e script shell que verifica status do repositório a ser executado
    strCMD = "C:\Program Files\Git\git-bash.exe C:\Fontes\ShellGit.sh"

'Executa o comando que irá gerar um log com o status do repositório
    Call ftaComandoShell(strCMD)
           
           
'Lendo o arquivo com o retorno do git status    '
    Dim strStatusComando As String
    
    Open "C:\Fontes\status_git.txt" For Input As #1
    Input #1, temp1

    strStatusComando = Mid(temp1, 17, 47)

    Close #1


    'strStatusComando = Replace(strStatusComando, "'", "")
    
'Se a branch estiver atualizada o status deve ser: Your branch is up to date
    'strSatusComando = Mid(strStatusComando, 1, 26)
    'MsgBox strSatusComando, vbInformation
    
'Se houver atualização o status deve ser: Your branch is behind
    strSatusComando = Mid(strStatusComando, 1, 26)
    MsgBox strSatusComando, vbInformation
    
    
'Pega o status para verificar se há atualizações (Your branch is up to date)
'    Dim strRecebe As String
'    strRecebe = Mid(CStr(strStatusComando), 1, 26)
    
 'Pega o status para verificar se há atualizações (Your branch is behind)
    Dim strRecebe As String
    strRecebe = Mid(CStr(strStatusComando), 1, 26)
    
    
    If ftaRegexEmail(strRecebe) = ftaRegexEmail("Your branch is behind") Then
                
        Select Case MsgBox("Foram encontradas novas atualizações para o sistema, deseja atualizar o sistema agora?", _
                            vbYesNo Or vbInformation Or vbDefaultButton1, "Atualização do sistema")
            
            Case vbYes
                MsgBox "O sistema será atualizado agora", vbInformation
                
                'Terminal e script shell que verifica status do git pull
                    strCMD = "C:\Program Files\Git\git-bash.exe C:\Fontes\ShellGitPull.sh"
                
                'Executa o comando que irá gerar um log com o status do repositório
                    Call ftaComandoShell(strCMD)
                
                
                'Lendo o arquivo com o retorno do git pull
                    Open "C:\Fontes\status_git_pull.txt" For Input As #2
                    Input #2, temp2
                
                        strStatusComando = temp2
                
                    Close #2
                    
                    
                    MsgBox strStatusComando, vbInformation
            Case vbNo
                MsgBox "Você não quis atualizar o sistema", vbInformation
                
        End Select
    End If

End Sub

'Call staImprimiTextoArquivo(strRecebe, "C:\Fontes\status_teste.txt")

'Função shell
Function ftaComandoShell(commandStr As String) As String
    
    Dim strRetorno As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set WshShell = CreateObject("WSCript.shell")
    strRetorno = WshShell.Exec(commandStr).StdOut.ReadAll


    ftaComandoShell = strRetorno
    
End Function

