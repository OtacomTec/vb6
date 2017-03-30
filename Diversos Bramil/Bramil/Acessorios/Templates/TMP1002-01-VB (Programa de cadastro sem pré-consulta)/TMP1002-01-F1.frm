VERSION 5.00
Begin VB.Form FormAlteraChave1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração do Código1"
   ClientHeight    =   1005
   ClientLeft      =   3375
   ClientTop       =   2625
   ClientWidth     =   3150
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   3150
   Begin VB.CommandButton CommandCancela 
      Caption         =   "&Cancela"
      Height          =   435
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   525
      Width           =   975
   End
   Begin VB.CommandButton CommandOk 
      Caption         =   "&Ok"
      Height          =   435
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   45
      Width           =   975
   End
   Begin VB.TextBox TextiNovoCodUsuario 
      Height          =   315
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   3
      Top             =   540
      Width           =   735
   End
   Begin VB.TextBox TextiAtualCodUsuario 
      Height          =   315
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   1
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Novo Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   165
      TabIndex        =   2
      Top             =   540
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código Atual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   1095
   End
End
Attribute VB_Name = "FormAlteraChave1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call ppCarregaPropriedadesForm(Me, 1001)
    
    TextiAtualCodUsuario = piCodigo1
    TextiAtualCodUsuario.Enabled = False
End Sub

Private Sub CommandCancela_Click()
    TextiNovoCodUsuario.Text = ""
    TextiNovoCodUsuario.SetFocus
End Sub

Private Sub CommandOk_Click()
    On Error GoTo Erro
    
    If TextiNovoCodUsuario.Text = "" Then
        Unload Me
        Exit Sub
    End If
    
    Dim pstrSql As String
    Dim prstTabela1 As Recordset
    pstrSql = "SELECT * FROM tUsuariosGrupo INNER JOIN tUsuarios ON tUsuariosGrupo.bCodGrupoUsuariotGrpUsu = tUsuarios.bCodGrupoUsuariotUsu WHERE tUsuarios.iCodUsuariotUsu = " & Val(TextiNovoCodUsuario)
    
    Set prstTabela1 = pdbConfus.OpenRecordset(pstrSql)
        
    If Not prstTabela1.EOF Then
        MsgBox "Código já utilizado para outro usuário", vbCritical, "Nome do Evento"
        prstTabela1.Close
        TextiNovoCodUsuario.SetFocus
        Exit Sub
    End If
    
    prstTabela1.Close
    FormCadastroUsuarios.TextiCodUsuariotUsu = TextiNovoCodUsuario
    Unload Me
                
    Exit Sub
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "CommandOk_Click"
    Err.Clear
End Sub
