VERSION 5.00
Begin VB.Form frmAdminLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   5130
   ClientTop       =   3930
   ClientWidth     =   5205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAdminLogin.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLogin 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2610
      Width           =   1875
   End
   Begin VB.TextBox txtLogin 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   " "
      Top             =   2040
      Width           =   1875
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Versão: 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4140
      TabIndex        =   4
      Top             =   2700
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   720
   End
End
Attribute VB_Name = "frmAdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conexao_login As New DLLConexao_Sistema.Conexao
Dim rstcomparacao As New ADODB.Recordset
Dim strsql As String
Dim strSenha As String
Dim strEmpresa As String
Dim strUsuario As String
Dim intNivel_usuario As Integer
Public intCodigo_usuario As Integer
Dim datValidade_usuario As Date
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
   End If
   If KeyAscii = 27 Then
        End
   End If
End Sub
Private Sub txtLogin_LostFocus(Index As Integer)
   If Trim(txtLogin(Index).Text) = "" Then Exit Sub
    
    Select Case Index
        Case 0 'Usuário
            On Error GoTo Erro
            
            strsql = "SELECT PKCodigo_TBUsuario,FKCodigo_TBEmpresa,DFNome_TBUsuario,DFSenha_TBUsuario,DFNivel_TBUsuario,IXData_validade_TBUsuario,DFPrazo_expira_senha_TBUsuario,IXData_cadastro_TBUsuario,DFProxima_troca_senha_TBUsuario FROM TBUsuario WHERE DFNome_TBUsuario = '" & txtLogin(0).Text & "'"
            conexao_login.Abrir_conexao ("Otica")
            
            Call Movimentacoes.Select_geral(strsql, "BDRetaguarda", rstcomparacao, "Otica", Me)
            
            'Setando e passando a estação local para a mensagem do intercomunicador
            Dim FCRegistro As DLLSystemManager.Registro
            Set FCRegistro = New DLLSystemManager.Registro
            strEstação = FCRegistro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName")
            
            If rstcomparacao.EOF = True And rstcomparacao.BOF = True Then
               MsgBox "Usuário não cadastrado", vbCritical, "Logicx"
               txtLogin(0).SetFocus
            Else
               Dim datValidade_senha As Date
               datValidade = Format(rstcomparacao!IXData_validade_TBUsuario, "DD/MM/YYYY")
               datValidade_senha = Format(rstcomparacao!DFProxima_troca_senha_TBUsuario, "DD/MM/YYYY")
               intCodigo_usuario = rstcomparacao!PKCodigo_TBUsuario
               strEmpresa = rstcomparacao!FKCodigo_TBEmpresa
               strSenha = rstcomparacao!DFSenha_TBUsuario
               intNivel_usuario = rstcomparacao!DFNivel_TBUsuario
               strUsuario = rstcomparacao!DFNome_TBUsuario
               intValidade_Senha = rstcomparacao!DFPrazo_expira_senha_TBUsuario
               'Verificando a validade da conta do usuário
               If datValidade <= CDate(Format(Now, "DD/MM/YYYY")) Then
                  MsgBox "Seu usuário não é válido, sua conta expirou em: " & Format(rstcomparacao!IXData_validade_TBUsuario, "DD/MM/YYYY") & " !Verifique com o administrador do Sistema.", vbInformation, "Logicx"
                  End
               End If
               'Verificando a senha do usuário
               If datValidade_senha <= CDate(Format(Now, "DD/MM/YYYY")) Then
                  intRetorno = MsgBox("Por motivos de segurança sua senha expirou (" & Format(datValidade_senha, "DD/MM/YYYY") & ")!,cadastre uma nova senha no sistema.", vbOKCancel, "Logicx")
                  'Sendo OK
                  If intRetorno = 1 Then
                     frmRotina_Troca_Senha.Show
                     frmRotina_Troca_Senha.strUsuario = strUsuario
                     frmRotina_Troca_Senha.strEstacao = strEstação
                     Unload Me
                     Exit Sub
                  Else
                     MsgBox "Enquanto não houver atualização de sua senha no sistema, vc estará impossibilitado de acessá-lô.", vbCritical, "Logicx"
                     End
                  End If
               End If
            End If
            
            Set rstcomparacao = Nothing
            conexao_login.Fechar_conexao
            Me.txtLogin(0).Text = UCase(Me.txtLogin(0).Text)
            
        Case 1 'Senha
            If Trim(txtLogin(0).Text) = "" Then
                txtLogin(0).SetFocus
                Exit Sub
            Else
                If strSenha = txtLogin(1).Text Then
                    ValidarUsuárioSenha = True
                    'Setando e passando a estação local para a mensagem do intercomunicador
            '        Dim FCRegistro As DLLSystemManager.Registro
            '        Set FCRegistro = New DLLSystemManager.Registro
            '        strEstação = FCRegistro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName")
                    'Ativando a manutenção do banco de Log ----> Movendo o mesmo para o Morto
                    If intNivel_usuario >= 9 Then
                       Dim rstParametros_log As New ADODB.Recordset
                       strsql = Empty
                       strsql = "SELECT * FROM TBParametros_Log"
                       Movimentacoes.Select_geral strsql, "BDLog", rstParametros_log, "Otica", Me
                       If CDate(rstParametros_log!DFProxima_Limpeza_Log) <= CDate(Format(Now, "DD/MM/YYYY")) Then
                          intRetorno = MsgBox("Solicita-se que de acordo com os parâmetros do sistema que se execute a rotina de limpeza do banco de log, prevista para: " & Format(rstParametros_log!DFProxima_Limpeza_Log, "DD/MM/YYYY") & ",deseja faze-lâ agora?", vbYesNo, "Logicx")
                          'Sendo Sim
                          If intRetorno = 6 Then
                             frmRotina_Limpeza_Log.Show
                             frmRotina_Limpeza_Log.strUsuario = strUsuario
                             frmRotina_Limpeza_Log.strEstacao = strEstação
                             Set rstParametros_log = Nothing
                             Unload Me
                            Exit Sub
                          End If
                       End If
                    End If
               End If
               If strSenha = txtLogin(1).Text Then
                    ValidarUsuárioSenha = True
                    frmAdminMDI.Show
                    'Adicionando um Novo Componente AplicativoUsuário
                    NovoLogin Trim(Me.txtLogin(0).Text), Trim(Me.txtLogin(1).Text), Str(intCodigo_usuario), strEmpresa, intNivel_usuario
                    'Gravando as inf do usuário no registro para contingência
                    Movimentacoes.Grava_Contingencia_Acessibilidade strEstação, txtLogin(0).Text, intCodigo_usuario, strEmpresa, "Otica"
                    Unload Me
                    KeyAscii = 0
                Else
                    MsgBox "Senha Inválida", vbCritical, "Logicx"
                    txtLogin(1).Text = ""
                    txtLogin(1).SetFocus
                End If
            End If
    End Select
    Exit Sub
    
Erro:
    If Err.Number = -2147220503 Then
       MsgBox "Fonte de dados não encontrada!", vbCritical, "Logicx"
       Shell App.Path & "Configurador de Sistemas.exe", vbNormalFocus
       End
       Exit Sub
    Else
       Call Erro.Erro(Me, "Otica", "Load", Err.Number)
       Exit Sub
    End If
    
End Sub
