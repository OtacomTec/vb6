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
   Begin VB.ComboBox cmbPDV 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmAdminLogin.frx":B747
      Left            =   2100
      List            =   "frmAdminLogin.frx":B749
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtSenha 
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
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2610
      Width           =   1875
   End
   Begin VB.TextBox txtOperador 
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
      Left            =   120
      TabIndex        =   0
      Text            =   " "
      Top             =   2040
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "N° PDV"
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
      Index           =   2
      Left            =   2145
      TabIndex        =   6
      Top             =   1800
      Width           =   660
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Operador:"
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
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "frmAdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String
Dim rstOperador As New ADODB.Recordset
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


Private Sub Form_Load()
    
    Dim rstPDV As New ADODB.Recordset
    
    'Carregando a combo de pdv
    strSql = Empty
    strSql = "SELECT PKCofigo_TBPdv FROM TBPDV"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstPDV, "Otica", Me
    
    'Montando a combo
    Do While rstPDV.EOF = False
        cmbPDV.AddItem rstPDV!PKCofigo_TBPdv
        rstPDV.MoveNext
    Loop
    
    Set rstPDV = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstOperador = Nothing
End Sub

Private Sub txtOperador_GotFocus()
'    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtOperador_LostFocus()

    On Error GoTo Erro
    
    Me.txtOperador.Text = UCase(Me.txtOperador)
    strSql = Empty
    strSql = "SELECT * FROM TBOperadores_ecf WHERE DFNome_TBOperadores_ecf = '" & Me.txtOperador.Text & "'"
    
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstOperador, "Otica", Me
  
    If rstOperador.BOF = True And rstOperador.EOF = True Then
       MsgBox "Operador não cadastrado.Verifique!", vbCritical, "Only Tech"
       txtOperador.Text = ""
       txtOperador.SetFocus
       Set rstOperador = Nothing
       Exit Sub
    End If
    
    Set rstOperador = Nothing
    
    'Me.txtSenha.SetFocus
    Me.cmbPDV.SetFocus
    
    Exit Sub
    
Erro:
    If Err.Number = -2147220503 Then
       MsgBox "Fonte de dados não encontrada!", vbCritical, "Only Tech"
       Shell App.Path & "Configurador de Sistemas.exe", vbNormalFocus
       End
       Exit Sub
    Else
       Call Erro.Erro(Me, "Otica", "Load", Err.Number)
       Exit Sub
    End If
    
End Sub

Private Sub txtSenha_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtSenha_LostFocus()

    On Error GoTo Erro
    
    If Me.txtSenha.Text <> "" Then
          strSql = Empty
          strSql = "SELECT * FROM TBOperadores_ecf WHERE DFSenha_TBOperadores_ecf = '" & Me.txtSenha.Text & "'"
          
          Movimentacoes.Select_geral strSql, "BDRetaguarda", rstOperador, "Otica", Me
        
          If rstOperador.BOF = True And rstOperador.EOF = True Then
             MsgBox "Senha não cadastrada.Verifique!", vbCritical, "Only Tech"
             Set rstOperador = Nothing
             txtSenha.Text = Empty
             txtSenha.SetFocus
             Exit Sub
          End If
          
          If Me.cmbPDV.Text = "" Then
             MsgBox "Favor indicar o número deste PDV!", vbCritical, "Only Tech"
             Set rstOperador = Nothing
             txtSenha.SetFocus
             Exit Sub
          End If
          
          'Verifcando se o operador já fechou susas operações neste PDV no dia de hj e verifica se há necessidade de se abrir o caixa
          Dim rstCaixa_Aberto As New ADODB.Recordset
          
          strSql = Empty
          strSql = "SELECT * FROM TBOperacao_caixa " & _
                   "WHERE FKCodigo_TBPdv = " & cmbPDV.Text & " " & _
                   "AND FKCodigo_TBOperadores_ecf = " & rstOperador!PKCodigo_TBOperadores_ecf & " " & _
                   "AND DFData_TBOperacao_caixa = '" & Format(Now, "YYYYMMDD") & "' " & _
                   "AND DFStatus_aberto_fechado_TBOperacao_caixa = 0"
                   
          Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCaixa_Aberto, "Otica", Me
          
          'Verifica se existe a necessidade de se abrir o caixa
          If rstCaixa_Aberto.BOF = True And rstCaixa_Aberto.EOF = True Then
             
             Dim rstCaixa_Fechado As New ADODB.Recordset
             
             'Verifica se o caixa já foi fechado por este operador
             strSql = Empty
             strSql = "SELECT * FROM TBOperacao_caixa " & _
                      "WHERE FKCodigo_TBPdv = " & cmbPDV.Text & " " & _
                      "AND FKCodigo_TBOperadores_ecf = " & rstOperador!PKCodigo_TBOperadores_ecf & " " & _
                      "AND DFData_TBOperacao_caixa = '" & Format(Now, "YYYYMMDD") & "' " & _
                      "AND DFStatus_aberto_fechado_TBOperacao_caixa = 1"
                   
             Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCaixa_Fechado, "Otica", Me
             
             If rstCaixa_Fechado.BOF = True And rstCaixa_Fechado.EOF = True Then
                'Enviando inf. para cadastro de abertura de caixa
                frmAbertura_Caixa.strPDV = cmbPDV.Text
                frmAbertura_Caixa.strNumero_PDV = cmbPDV.Text
                frmAbertura_Caixa.strEmpresa_Operador = rstOperador!FKCodigo_TBEmpresa
                frmAbertura_Caixa.strCodigo_Operador = rstOperador!PKCodigo_TBOperadores_ecf
                frmAbertura_Caixa.strOperador = rstOperador!PKCodigo_TBOperadores_ecf & "-" & rstOperador!DFNome_TBOperadores_ecf
                frmAbertura_Caixa.Show
                Set rstCaixa_Aberto = Nothing
                Set rstCaixa_Fechado = Nothing
                Set rstOperador = Nothing
                Unload Me
             Else
                MsgBox "Este PDV já foi aberto e encerrado na data de hoje por este operador!Verifique.", vbCritical, "Only Tech"
                Set rstCaixa_Fechado = Nothing
                Set rstCaixa_Aberto = Nothing
                Set rstOperador = Nothing
                Me.txtOperador.Text = Empty
                Me.txtSenha.Text = Empty
                Me.txtOperador.SetFocus
                Exit Sub
             End If
             
             Set rstCaixa_Fechado = Nothing
             
          Else
            'Verifica se o caixa já foi fechado por este operador
            strSql = Empty
            strSql = "SELECT * FROM TBOperacao_caixa " & _
                     "WHERE FKCodigo_TBPdv = " & cmbPDV.Text & " " & _
                     "AND FKCodigo_TBOperadores_ecf = " & rstOperador!PKCodigo_TBOperadores_ecf & " " & _
                     "AND DFData_TBOperacao_caixa = '" & Format(Now, "YYYYMMDD") & "' " & _
                     "AND DFStatus_aberto_fechado_TBOperacao_caixa = 1"
                     
            Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCaixa_Fechado, "Otica", Me
             
            If Not rstCaixa_Fechado.BOF = True And Not rstCaixa_Fechado.EOF = True Then
                MsgBox "Este PDV já foi aberto e encerrado na data de hoje por este operador!Verifique.", vbCritical, "Only Tech"
                Set rstCaixa_Fechado = Nothing
                Set rstCaixa_Aberto = Nothing
                Set rstOperador = Nothing
                Me.txtOperador.Text = Empty
                Me.txtSenha.Text = Empty
                Me.txtOperador.SetFocus
                Exit Sub
            End If
          
          
            'Enviando inf. para Tela de Vendas
            frmTela_Venda.strEmpresa_Operador = rstOperador!FKCodigo_TBEmpresa
            frmTela_Venda.strOperador = rstOperador!PKCodigo_TBOperadores_ecf & "-" & rstOperador!DFNome_TBOperadores_ecf
            frmTela_Venda.strCodigo_Operador = rstOperador!PKCodigo_TBOperadores_ecf
            frmTela_Venda.strPDV = Me.cmbPDV.Text
            frmTela_Venda.Show
            Set rstCaixa_Aberto = Nothing
            Set rstOperador = Nothing
            Unload Me
          End If
          
          Set rstCaixa_Aberto = Nothing
          Set rstOperador = Nothing
          Unload Me
    End If
    
    Exit Sub
    
Erro:
    If Err.Number = -2147220503 Then
       MsgBox "Fonte de dados não encontrada!", vbCritical, "Only Tech"
       Shell App.Path & "Configurador de Sistemas.exe", vbNormalFocus
       End
       Exit Sub
    Else
       Call Erro.Erro(Me, "Otica", "Load", Err.Number)
       Exit Sub
    End If
    
End Sub
