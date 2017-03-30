Attribute VB_Name = "Module1"
Public SvMsg As VetorDeMensagens.ServidorDeMensagens
Public AT As AreaDeTrabalho
Public FCRegistro As DLLFuncoesGerais.Registro
Public strEsta��o As String
Public strEsta��oComent�rio As String

Public strControleDeUsu�rios As String 'Controlar todos os usu�rios Logados separados por |
'    pstrComputer = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName")
'    pstrDescrEstacao = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\VxD\VNETSUP\", "Comment")


Public Function ExibirLoginOp��es(Optional FormAtivo As Boolean = False)
    Dim I As Integer
    Dim mtzUsu�rios() As String
    
    mtzUsu�rios = Split(frmAdminMDI.AplicativoUsu�rio(0).Tag, "|")
    frmAdminLoginOpcoes.lstUsu�riosLogados.Clear
    For I = 0 To UBound(mtzUsu�rios) - 1
    'For i = 0 To frmAdminMDI.AplicativoUsu�rio.ubound
        'If i = 0 Then
        If mtzUsu�rios(I) = 0 Then
            frmAdminLoginOpcoes.lstUsu�riosLogados.AddItem frmAdminMDI.AplicativoUsu�rio(I).NomeReduzido & " (Atual)"
        Else
            'If Not frmAdminMDI.AplicativoUsu�rio(i).Janela Is Nothing Then
            'If frmAdminMDI.AplicativoUsu�rio(i) Is Nothing = False Then
                frmAdminLoginOpcoes.lstUsu�riosLogados.AddItem frmAdminMDI.AplicativoUsu�rio(mtzUsu�rios(I)).NomeReduzido
                'frmAdminMDI.AplicativoUsu�rio.Item.
            'End If
        End If
    Next I
    If FormAtivo = False Then frmAdminLoginOpcoes.Show 1
End Function

Public Function NovoLogin(strLogin As String, strSenha As String)
    Dim I As Integer
    Dim mtzUsu�riosLogados() As String

    'Verificando Janelas abertas
    mtzUsu�riosLogados = Split(frmAdminMDI.AplicativoUsu�rio(0).Tag, "|")
    For I = 1 To UBound(mtzUsu�riosLogados) - 1
        If Not frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela Is Nothing = True Then
            If frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela.Usu�rio = strLogin Then
                'Habilita a �rea de Trabalho quando houver outra j� aberta
                frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela.Enabled = True
            Else
                frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela.Enabled = False
                frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela.WindowState = 1
            End If
        End If
    Next I
    
    
    Dim frm As frmAdminDeskTopCliente
    Set frm = New frmAdminDeskTopCliente
    
    I = frmAdminMDI.AplicativoUsu�rio.UBound + 1
    Load frmAdminMDI.AplicativoUsu�rio(I)
    frmAdminMDI.AplicativoUsu�rio(I).Nome = strLogin
    frmAdminMDI.AplicativoUsu�rio(I).NomeReduzido = strLogin
    frmAdminMDI.AplicativoUsu�rio(I).Senha = strSenha
    
    'Criando uma �rea de Trabalho
    frmAdminMDI.AplicativoUsu�rio(I).�reaDeTrabalho = AT.Criar�reaDeTrabalho(frm)
    Set frmAdminMDI.AplicativoUsu�rio(I).Janela = frm
    frmAdminMDI.AplicativoUsu�rio(I).Janela.ID_Usu�rio = I
    
    frm.Caption = frmAdminMDI.AplicativoUsu�rio(I).NomeReduzido & " / �rea de Trabalho"
    'frm.Show
    frm.Usu�rio = frmAdminMDI.AplicativoUsu�rio(I).NomeReduzido
    frm.Senha = frmAdminMDI.AplicativoUsu�rio(I).Senha
    frm.ID_Usu�rio = I
                                        
    frmAdminMDI.AplicativoUsu�rio(0).Nome = frmAdminMDI.AplicativoUsu�rio(I).Nome
    frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = frmAdminMDI.AplicativoUsu�rio(I).NomeReduzido
    frmAdminMDI.AplicativoUsu�rio(0).Senha = frmAdminMDI.AplicativoUsu�rio(I).Senha
    frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho = frmAdminMDI.AplicativoUsu�rio(I).�reaDeTrabalho
    
    'Armazenar� todos os Usu�rios Logados no Sistema (adicionando-os)
    If frmAdminMDI.AplicativoUsu�rio(0).Tag = Empty Then frmAdminMDI.AplicativoUsu�rio(0).Tag = "0|"
    frmAdminMDI.AplicativoUsu�rio(0).Tag = frmAdminMDI.AplicativoUsu�rio(0).Tag & I & "|"
    Set frmAdminMDI.AplicativoUsu�rio(0).Janela = frmAdminMDI.AplicativoUsu�rio(I).Janela
    frmAdminMDI.Arrange 3
    
End Function
Public Function ReativarLogin(strLogin As String, strSenha As String) As Boolean
    
    Dim strsql As String
    Dim rstComparacao_senha As New ADODB.Recordset
    Dim conexao_senha As New DLLConexao_Sistema.Conexao
    
    If frmAdminMDI.AplicativoUsu�rio.UBound = 0 Then
        MsgBox "N�o h� Usu�rios Logados"
        Exit Function
    End If
    
    strsql = "SELECT DFSenha_TBUsuario FROM TBUsuario WHERE DFNome_TBUsuario = '" & strLogin & "' "
    conexao_senha.Abrir_conexao ("PDV")
    
    Call Movimentacoes.Select_geral(strsql, "BDSupervisor", rstComparacao_senha)
    
    If rstComparacao_senha!DFSenha_TBUsuario = strSenha Then
          ValidarUsu�rioSenhaJ�Logado = True
          If ValidarUsu�rioSenhaJ�Logado = True Then
              If frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = strLogin Then
                  MSG = MSG & "O Usu�rio " & strLogin & " j� � o usu�rio Ativo!"
                  MsgBox MSG
                  If frmAdminMDI.AplicativoUsu�rio(0).Janela.Enabled = False Then frmAdminMDI.AplicativoUsu�rio(0).Janela.Enabled = True
                  frmAdminMDI.AplicativoUsu�rio(0).Janela.WindowState = 2
                  'If frmAdminMDI.ActiveForm.Enabled = False Then frmAdminMDI.ActiveForm.Enabled = True
                  frmAdminMDI.Arrange 3
                  Exit Function
              End If
              
              'Minimizo a AT do Usu�rio Atual
              If Not frmAdminMDI.AplicativoUsu�rio(0).Janela Is Nothing = True Then
                  frmAdminMDI.AplicativoUsu�rio(0).Janela.Enabled = False
                  frmAdminMDI.AplicativoUsu�rio(0).Janela.WindowState = 1
                  frmAdminMDI.AplicativoUsu�rio(0).Nome = Empty
                  frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = Empty
                  frmAdminMDI.AplicativoUsu�rio(0).Senha = Empty
                  Set frmAdminMDI.AplicativoUsu�rio(0).Janela = Nothing
                  frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho = Empty
              End If
              
              Achou = False
              'Inicio procura do Usu�rio (j� logado) a ser Ativado
              mtzUsu�riosLogados = Split(frmAdminMDI.AplicativoUsu�rio(0).Tag, "|")
              For I = 1 To UBound(mtzUsu�riosLogados) - 1
                  'Se a Janela(mtzUsu�riosLogados(i) n�o for vazia ent�o Faz
                  If Not frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)) Is Nothing = True Then
                      If frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).NomeReduzido = strLogin Then
                          frmAdminMDI.AplicativoUsu�rio(0).Nome = frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Nome
                          frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).NomeReduzido
                          frmAdminMDI.AplicativoUsu�rio(0).Senha = frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Senha
                          Set frmAdminMDI.AplicativoUsu�rio(0).Janela = frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela
                          frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho = frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).�reaDeTrabalho
                          
                          frmAdminMDI.AplicativoUsu�rio(0).Janela.WindowState = 2
                          
                          If frmAdminMDI.AplicativoUsu�rio(0).Janela.Enabled = False Then frmAdminMDI.AplicativoUsu�rio(0).Janela.Enabled = True
                          frmAdminMDI.AplicativoUsu�rio(0).Janela.ID_Usu�rio = mtzUsu�riosLogados(I)
                          Achou = True
                          ReativarLogin = True
                      Else
                          If frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela.Enabled = True Then frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela.Enabled = False
                          If frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela.WindowState <> 1 Then frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela.WindowState = 1
                      End If
                  End If
              Next I
          End If
          If Achou = False Then
              MsgBox "O usu�rio " & strLogin & " n�o pode ser Ativado!" & Chr(13) & "N�o h� Usu�rio Atual."
          End If
          ExibirLoginOp��es True
          frmAdminMDI.Arrange 3
    Else
        MsgBox "Senha n�o confere!", vbCritical, "Logicx"
    End If
    
    Set rstComparacao_senha = Nothing
    conexao_senha.Fechar_conexao
    
End Function

Public Function LogOff(strLogin As String, strSenha As String)

    Dim mtzUsu�riosLogados() As String
    Dim booAchou As Boolean
    Dim strsql As String
    Dim rstComparacao_senha As New ADODB.Recordset
    Dim conexao_senha As New DLLConexao_Sistema.Conexao
    
    strsql = "SELECT DFSenha_TBUsuario FROM TBUsuario WHERE DFNome_TBUsuario = '" & strLogin & "' "
    conexao_senha.Abrir_conexao ("PDV")
    
    Call Movimentacoes.Select_geral(strsql, "BDSupervisor", rstComparacao_senha)
    
    If rstComparacao_senha!DFSenha_TBUsuario = strSenha Then
        If frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = strLogin Then
            frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = Empty
            frmAdminMDI.AplicativoUsu�rio(0).Nome = Empty
            frmAdminMDI.AplicativoUsu�rio(0).Senha = Empty
            frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho = Empty
            Set frmAdminMDI.AplicativoUsu�rio(0).Janela = Nothing
        End If
                
        mtzUsu�riosLogados = Split(frmAdminMDI.AplicativoUsu�rio(0).Tag, "|")
        For I = 1 To UBound(mtzUsu�riosLogados) - 1
            If frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela.Usu�rio = strLogin Then
                Unload frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela
                Unload frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I))
                frmAdminMDI.AplicativoUsu�rio(0).Tag = Replace(frmAdminMDI.AplicativoUsu�rio(0).Tag, "|" & mtzUsu�riosLogados(I) & "|", "|")
                booAchou = True
            End If
        Next I
        If booAchou = False Then MsgBox "O Usu�rio " & strLogin & " n�o est� Logado!"
    Else
        MsgBox "Senha n�o confere!", vbCritical, "Logicx"
    End If
    
    Set rstComparacao_senha = Nothing
    conexao_senha.Fechar_conexao
    
End Function

Public Function DesativarLogin(strLogin As String, strSenha As String)

    Dim mtzUsu�riosLogados() As String
    Dim booAchou As Boolean
    Dim strsql As String
    Dim rstComparacao_senha As New ADODB.Recordset
    Dim conexao_senha As New DLLConexao_Sistema.Conexao
    
    strsql = "SELECT DFSenha_TBUsuario FROM TBUsuario WHERE DFNome_TBUsuario = '" & strLogin & "' "
    conexao_senha.Abrir_conexao ("PDV")
    
    Call Movimentacoes.Select_geral(strsql, "BDSupervisor", rstComparacao_senha)
    
    If rstComparacao_senha!DFSenha_TBUsuario = strSenha Then
        If frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = strLogin Then
            frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = Empty
            frmAdminMDI.AplicativoUsu�rio(0).Nome = Empty
            frmAdminMDI.AplicativoUsu�rio(0).Senha = Empty
            frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho = Empty
            Set frmAdminMDI.AplicativoUsu�rio(0).Janela = Nothing
            ExibirLoginOp��es True
        Else
            mtzUsu�riosLogados = Split(frmAdminMDI.AplicativoUsu�rio(0).Tag, "|")
            For I = 1 To UBound(mtzUsu�riosLogados) - 1
                If frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela.Usu�rio = strLogin Then
                    booAchou = True
                    Exit For
                End If
            Next I
            If booAchou Then
                MsgBox "O Usu�rio " & strLogin & " j� � um usu�rio Logado Inativo!"
            Else
                MsgBox "O Usu�rio " & strLogin & " n�o est� Logado!"
            End If
        End If
    Else
        MsgBox "Senha n�o confere!", vbCritical, "Logicx"
    End If
    
    Set rstComparacao_senha = Nothing
    conexao_senha.Fechar_conexao
    
End Function


Public Function Fechar�reaDeTrabalho(frm As Form)
    Dim mtzUsu�riosLogados() As String
    Dim booAchou As Boolean
    
    Unload frmAdminMDI.AplicativoUsu�rio(frm.ID_Usu�rio)
    frmAdminMDI.AplicativoUsu�rio(0).Nome = Empty
    frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = Empty
    frmAdminMDI.AplicativoUsu�rio(0).Senha = Empty
    frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho = Empty
    
    'Retira o ID do Usu�rio da Tag que controla os Usu�rios Logados
    'A Tag utilizada � do componente OCX AplicativoUsu�rio(0) que � respons�vel
    'por armazenar dados do Usu�rio Atual.
    frmAdminMDI.AplicativoUsu�rio(0).Tag = Replace(frmAdminMDI.AplicativoUsu�rio(0).Tag, "|" & frm.ID_Usu�rio & "|", "|")
    Set frmAdminMDI.AplicativoUsu�rio(0).Janela = Nothing
            
    'Verificar se h� outras �reas de Trabalho desse Usu�rio, Avis�-lo se houver e Ativar a �rea de Trabalho
    mtzUsu�riosLogados = Split(frmAdminMDI.AplicativoUsu�rio(0).Tag, "|")
    For I = 1 To UBound(mtzUsu�riosLogados) - 1
        If frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela.Usu�rio = frm.Usu�rio Then
            
            frmAdminMDI.AplicativoUsu�rio(0).Nome = frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Nome
            frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).NomeReduzido
            frmAdminMDI.AplicativoUsu�rio(0).Senha = frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Senha
            frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho = frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).�reaDeTrabalho
            Set frmAdminMDI.AplicativoUsu�rio(0).Janela = frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela
            MSG = MSG & "O Sistema detectou outras �reas de Trabalho com o seu Login!" & Chr(13) & "Feche-as se for necess�rio." & Chr(13)
            MSG = MSG & "O Sistema exibir� agora a �rea de Trabalho encontrada. "
            MsgBox MSG
            frmAdminMDI.AplicativoUsu�rio(mtzUsu�riosLogados(I)).Janela.WindowState = 2
            Exit For
        End If
    Next I
End Function
