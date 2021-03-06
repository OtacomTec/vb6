Attribute VB_Name = "Module1"
Public AT As AreaDeTrabalho
Public FCRegistro As DLLSystemManager.Registro
Public strEstaçăo As String
Public strEstaçăoComentário As String

Public strControleDeUsuários As String 'Controlar todos os usuários Logados separados por |
'    pstrComputer = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName")
'    pstrDescrEstacao = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\VxD\VNETSUP\", "Comment")


Public Function ExibirLoginOpçőes(Optional FormAtivo As Boolean = False)
    Dim I As Integer
    Dim mtzUsuários() As String
    
    mtzUsuários = Split(frmAdminMDI.AplicativoUsuário(0).Tag, "|")
    frmAdminLoginOpcoes.lstUsuáriosLogados.Clear
    For I = 0 To UBound(mtzUsuários) - 1
    'For i = 0 To frmAdminMDI.AplicativoUsuário.ubound
        'If i = 0 Then
        If mtzUsuários(I) = 0 Then
            frmAdminLoginOpcoes.lstUsuáriosLogados.AddItem frmAdminMDI.AplicativoUsuário(I).NomeReduzido & " (Atual)"
        Else
            'If Not frmAdminMDI.AplicativoUsuário(i).Janela Is Nothing Then
            'If frmAdminMDI.AplicativoUsuário(i) Is Nothing = False Then
                frmAdminLoginOpcoes.lstUsuáriosLogados.AddItem frmAdminMDI.AplicativoUsuário(mtzUsuários(I)).NomeReduzido
                'frmAdminMDI.AplicativoUsuário.Item.
            'End If
        End If
    Next I
    If FormAtivo = False Then frmAdminLoginOpcoes.Show 1
End Function

Public Function NovoLogin(strLogin As String, strSenha As String, strCodigo_usuario As String, strCodigo_Empresa As String, intNivel_user As Integer)
    
    Dim I As Integer
    Dim mtzUsuáriosLogados() As String

    'Verificando Janelas abertas
    mtzUsuáriosLogados = Split(frmAdminMDI.AplicativoUsuário(0).Tag, "|")
    For I = 1 To UBound(mtzUsuáriosLogados) - 1
        If Not frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela Is Nothing = True Then
            If frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela.Usuário = strLogin Then
                'Habilita a Área de Trabalho quando houver outra já aberta
                frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela.Enabled = True
            Else
                frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela.Enabled = False
                frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela.WindowState = 1
            End If
        End If
    Next I
    
    Dim frm As frmAdminDeskTopCliente
    Set frm = New frmAdminDeskTopCliente
    
    I = frmAdminMDI.AplicativoUsuário.UBound + 1
    Load frmAdminMDI.AplicativoUsuário(I)
    frmAdminMDI.AplicativoUsuário(I).Nome = strLogin
    frmAdminMDI.AplicativoUsuário(I).NomeReduzido = strLogin
    frmAdminMDI.AplicativoUsuário(I).Senha = strSenha
    frmAdminMDI.AplicativoUsuário(I).Codigo = strCodigo_usuario
    frmAdminMDI.AplicativoUsuário(I).Empresa = strCodigo_Empresa
    frmAdminMDI.AplicativoUsuário(I).Nível = intNivel_user
    
    'Passando o Id da area de trabalho para o Admin
    frmAdminMDI.intAtual_ID = I
    
    'Criando uma Área de Trabalho
    frmAdminMDI.AplicativoUsuário(I).ÁreaDeTrabalho = AT.CriarÁreaDeTrabalho(frm)
    Set frmAdminMDI.AplicativoUsuário(I).Janela = frm
    frmAdminMDI.AplicativoUsuário(I).Janela.ID_Usuário = I
    
    frm.Caption = frmAdminMDI.AplicativoUsuário(I).NomeReduzido & " / Área de Trabalho"
    
    frm.Usuário = frmAdminMDI.AplicativoUsuário(I).NomeReduzido
    frm.Senha = frmAdminMDI.AplicativoUsuário(I).Senha
    frm.ID_Usuário = I
                                        
    frmAdminMDI.AplicativoUsuário(0).Nome = frmAdminMDI.AplicativoUsuário(I).Nome
    frmAdminMDI.AplicativoUsuário(0).NomeReduzido = frmAdminMDI.AplicativoUsuário(I).NomeReduzido
    frmAdminMDI.AplicativoUsuário(0).Senha = frmAdminMDI.AplicativoUsuário(I).Senha
    frmAdminMDI.AplicativoUsuário(0).Codigo = frmAdminMDI.AplicativoUsuário(I).Codigo
    frmAdminMDI.AplicativoUsuário(0).Empresa = frmAdminMDI.AplicativoUsuário(I).Empresa
    frmAdminMDI.AplicativoUsuário(0).Nível = frmAdminMDI.AplicativoUsuário(I).Nível
    frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho = frmAdminMDI.AplicativoUsuário(I).ÁreaDeTrabalho
    
    'Armazenará todos os Usuários Logados no Sistema (adicionando-os)
    If frmAdminMDI.AplicativoUsuário(0).Tag = Empty Then frmAdminMDI.AplicativoUsuário(0).Tag = "0|"
    frmAdminMDI.AplicativoUsuário(0).Tag = frmAdminMDI.AplicativoUsuário(0).Tag & I & "|"
    Set frmAdminMDI.AplicativoUsuário(0).Janela = frmAdminMDI.AplicativoUsuário(I).Janela
    frmAdminMDI.Arrange 3
    
End Function
Public Function ReativarLogin(strLogin As String, strSenha As String, Aplicacao As String, Form As Form) As Boolean
    
    Dim strSql As String
    Dim rstComparacao_senha As New adodb.Recordset
    Dim conexao_senha As New DLLConexao_Sistema.Conexao
    
    If frmAdminMDI.AplicativoUsuário.UBound = 0 Then
        MsgBox "Năo há Usuários Logados"
        Exit Function
    End If
    
    strSql = "SELECT DFSenha_TBUsuario FROM TBUsuario WHERE DFNome_TBUsuario = '" & strLogin & "' "
    conexao_senha.Abrir_conexao ("Otica")
    
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstComparacao_senha, Aplicacao, Form)
    
    If rstComparacao_senha!DFSenha_TBUsuario = strSenha Then
          ValidarUsuárioSenhaJáLogado = True
          If ValidarUsuárioSenhaJáLogado = True Then
              If frmAdminMDI.AplicativoUsuário(0).NomeReduzido = strLogin Then
                  MSG = MSG & "O Usuário " & strLogin & " já é o usuário Ativo!"
                  MsgBox MSG
                  If frmAdminMDI.AplicativoUsuário(0).Janela.Enabled = False Then frmAdminMDI.AplicativoUsuário(0).Janela.Enabled = True
                  frmAdminMDI.AplicativoUsuário(0).Janela.WindowState = 2
                  'If frmAdminMDI.ActiveForm.Enabled = False Then frmAdminMDI.ActiveForm.Enabled = True
                  frmAdminMDI.Arrange 3
                  Exit Function
              End If
              
              'Minimizo a AT do Usuário Atual
              If Not frmAdminMDI.AplicativoUsuário(0).Janela Is Nothing = True Then
                  frmAdminMDI.AplicativoUsuário(0).Janela.Enabled = False
                  frmAdminMDI.AplicativoUsuário(0).Janela.WindowState = 1
                  frmAdminMDI.AplicativoUsuário(0).Nome = Empty
                  frmAdminMDI.AplicativoUsuário(0).NomeReduzido = Empty
                  frmAdminMDI.AplicativoUsuário(0).Senha = Empty
                  frmAdminMDI.AplicativoUsuário(0).Empresa = Empty
                  frmAdminMDI.AplicativoUsuário(0).Nível = Empty
                  frmAdminMDI.AplicativoUsuário(0).Codigo = Empty
                  Set frmAdminMDI.AplicativoUsuário(0).Janela = Nothing
                  frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho = Empty
              End If
              
              Achou = False
              'Inicio procura do Usuário (já logado) a ser Ativado
              mtzUsuáriosLogados = Split(frmAdminMDI.AplicativoUsuário(0).Tag, "|")
              For I = 1 To UBound(mtzUsuáriosLogados) - 1
                  'Se a Janela(mtzUsuáriosLogados(i) năo for vazia entăo Faz
                  If Not frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)) Is Nothing = True Then
                      If frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).NomeReduzido = strLogin Then
                          frmAdminMDI.AplicativoUsuário(0).Nome = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Nome
                          frmAdminMDI.AplicativoUsuário(0).NomeReduzido = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).NomeReduzido
                          frmAdminMDI.AplicativoUsuário(0).Senha = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Senha
                          frmAdminMDI.AplicativoUsuário(0).Codigo = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Codigo
                          frmAdminMDI.AplicativoUsuário(0).Empresa = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Empresa
                          frmAdminMDI.AplicativoUsuário(0).Nível = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Nível
                          Set frmAdminMDI.AplicativoUsuário(0).Janela = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela
                          frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).ÁreaDeTrabalho
                          
                          frmAdminMDI.AplicativoUsuário(0).Janela.WindowState = 2
                          
                          If frmAdminMDI.AplicativoUsuário(0).Janela.Enabled = False Then frmAdminMDI.AplicativoUsuário(0).Janela.Enabled = True
                          frmAdminMDI.AplicativoUsuário(0).Janela.ID_Usuário = mtzUsuáriosLogados(I)
                          Achou = True
                          ReativarLogin = True
                      Else
                          If frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela.Enabled = True Then frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela.Enabled = False
                          If frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela.WindowState <> 1 Then frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela.WindowState = 1
                      End If
                  End If
              Next I
          End If
          If Achou = False Then
              MsgBox "O usuário " & strLogin & " năo pode ser Ativado!" & Chr(13) & "Năo há Usuário Atual."
          End If
          ExibirLoginOpçőes True
          frmAdminMDI.Arrange 3
    Else
        MsgBox "Senha năo confere!", vbCritical, "Logicx"
    End If
    
    Set rstComparacao_senha = Nothing
    conexao_senha.Fechar_conexao
    
End Function

Public Function LogOff(strLogin As String, strSenha As String, Aplicacao As String, Form As Form)

    Dim mtzUsuáriosLogados() As String
    Dim booAchou As Boolean
    Dim strSql As String
    Dim rstComparacao_senha As New adodb.Recordset
    Dim conexao_senha As New DLLConexao_Sistema.Conexao
    
    strSql = "SELECT DFSenha_TBUsuario FROM TBUsuario WHERE DFNome_TBUsuario = '" & strLogin & "' "
    conexao_senha.Abrir_conexao ("Otica")
    
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstComparacao_senha, Aplicacao, Form)
    
    If rstComparacao_senha!DFSenha_TBUsuario = strSenha Then
        If frmAdminMDI.AplicativoUsuário(0).NomeReduzido = strLogin Then
            frmAdminMDI.AplicativoUsuário(0).NomeReduzido = Empty
            frmAdminMDI.AplicativoUsuário(0).Nome = Empty
            frmAdminMDI.AplicativoUsuário(0).Senha = Empty
            frmAdminMDI.AplicativoUsuário(0).Codigo = Empty
            frmAdminMDI.AplicativoUsuário(0).Empresa = Empty
            frmAdminMDI.AplicativoUsuário(0).Nível = Empty
            frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho = Empty
            Set frmAdminMDI.AplicativoUsuário(0).Janela = Nothing
        End If
                
        mtzUsuáriosLogados = Split(frmAdminMDI.AplicativoUsuário(0).Tag, "|")
        For I = 1 To UBound(mtzUsuáriosLogados) - 1
            If frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela.Usuário = strLogin Then
                Unload frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela
                Unload frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I))
                frmAdminMDI.AplicativoUsuário(0).Tag = Replace(frmAdminMDI.AplicativoUsuário(0).Tag, "|" & mtzUsuáriosLogados(I) & "|", "|")
                booAchou = True
            End If
        Next I
        If booAchou = False Then MsgBox "O Usuário " & strLogin & " năo está Logado!"
    Else
        MsgBox "Senha năo confere!", vbCritical, "Logicx"
    End If
    
    Set rstComparacao_senha = Nothing
    conexao_senha.Fechar_conexao
    
End Function

Public Function DesativarLogin(strLogin As String, strSenha As String, Aplicacao As String, Form As Form)

    Dim mtzUsuáriosLogados() As String
    Dim booAchou As Boolean
    Dim strSql As String
    Dim rstComparacao_senha As New adodb.Recordset
    Dim conexao_senha As New DLLConexao_Sistema.Conexao
    
    strSql = "SELECT DFSenha_TBUsuario FROM TBUsuario WHERE DFNome_TBUsuario = '" & strLogin & "' "
    conexao_senha.Abrir_conexao ("Otica")
    
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstComparacao_senha, Aplicacao, Form)
    
    If rstComparacao_senha!DFSenha_TBUsuario = strSenha Then
        If frmAdminMDI.AplicativoUsuário(0).NomeReduzido = strLogin Then
            frmAdminMDI.AplicativoUsuário(0).NomeReduzido = Empty
            frmAdminMDI.AplicativoUsuário(0).Nome = Empty
            frmAdminMDI.AplicativoUsuário(0).Senha = Empty
            frmAdminMDI.AplicativoUsuário(0).Codigo = Empty
            frmAdminMDI.AplicativoUsuário(0).Empresa = Empty
            frmAdminMDI.AplicativoUsuário(0).Nível = Empty
            frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho = Empty
            Set frmAdminMDI.AplicativoUsuário(0).Janela = Nothing
            ExibirLoginOpçőes True
        Else
            mtzUsuáriosLogados = Split(frmAdminMDI.AplicativoUsuário(0).Tag, "|")
            For I = 1 To UBound(mtzUsuáriosLogados) - 1
                If frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela.Usuário = strLogin Then
                    booAchou = True
                    Exit For
                End If
            Next I
            If booAchou Then
                MsgBox "O Usuário " & strLogin & " já é um usuário Logado Inativo!"
            Else
                MsgBox "O Usuário " & strLogin & " năo está Logado!"
            End If
        End If
    Else
        MsgBox "Senha năo confere!", vbCritical, "Logicx"
    End If
    
    Set rstComparacao_senha = Nothing
    conexao_senha.Fechar_conexao
    
End Function


Public Function FecharÁreaDeTrabalho(frm As Form)
    Dim mtzUsuáriosLogados() As String
    Dim booAchou As Boolean
    
    Unload frmAdminMDI.AplicativoUsuário(frm.ID_Usuário)
    frmAdminMDI.AplicativoUsuário(0).Nome = Empty
    frmAdminMDI.AplicativoUsuário(0).NomeReduzido = Empty
    frmAdminMDI.AplicativoUsuário(0).Senha = Empty
    frmAdminMDI.AplicativoUsuário(0).Codigo = Empty
    frmAdminMDI.AplicativoUsuário(0).Empresa = Empty
    frmAdminMDI.AplicativoUsuário(0).Nível = Empty
    frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho = Empty
    
    'Retira o ID do Usuário da Tag que controla os Usuários Logados
    'A Tag utilizada é do componente OCX AplicativoUsuário(0) que é responsável
    'por armazenar dados do Usuário Atual.
    frmAdminMDI.AplicativoUsuário(0).Tag = Replace(frmAdminMDI.AplicativoUsuário(0).Tag, "|" & frm.ID_Usuário & "|", "|")
    Set frmAdminMDI.AplicativoUsuário(0).Janela = Nothing
            
    'Verificar se há outras Áreas de Trabalho desse Usuário, Avisá-lo se houver e Ativar a Área de Trabalho
    mtzUsuáriosLogados = Split(frmAdminMDI.AplicativoUsuário(0).Tag, "|")
    For I = 1 To UBound(mtzUsuáriosLogados) - 1
        If frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela.Usuário = frm.Usuário Then
            
            frmAdminMDI.AplicativoUsuário(0).Nome = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Nome
            frmAdminMDI.AplicativoUsuário(0).NomeReduzido = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).NomeReduzido
            frmAdminMDI.AplicativoUsuário(0).Senha = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Senha
            frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).ÁreaDeTrabalho
            Set frmAdminMDI.AplicativoUsuário(0).Janela = frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela
            MSG = MSG & "O Sistema detectou outras Áreas de Trabalho com o seu Login!" & Chr(13) & "Feche-as se for necessário." & Chr(13)
            MSG = MSG & "O Sistema exibirá agora a Área de Trabalho encontrada. "
            MsgBox MSG
            frmAdminMDI.AplicativoUsuário(mtzUsuáriosLogados(I)).Janela.WindowState = 2
            Exit For
        End If
    Next I
End Function
