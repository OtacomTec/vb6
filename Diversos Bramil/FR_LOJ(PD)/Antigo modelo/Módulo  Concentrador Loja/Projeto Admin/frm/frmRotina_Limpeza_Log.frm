VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmRotina_Limpeza_Log 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotina de Limpeza do Log do Sistema"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRotina_Limpeza_Log.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pgbAndamento 
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   1620
      Visible         =   0   'False
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblLog_Morto 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   1050
      Width           =   60
   End
   Begin VB.Label lblTotal_Log 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   420
      Width           =   60
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total de Reg. Banco de LOG Morto:"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   780
      Width           =   3030
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total de Registros no Banco de LOG:"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   3150
   End
End
Attribute VB_Name = "frmRotina_Limpeza_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Ventura                                                        '
' Módulo.................: Admin                                                          '
' Objetivo...............: MDI Principal                                                  '
' Data de Criação........: 22/07/2004                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião,Rafael Gomes, Sérgio   '
' Última Manutenção......:                                                                '
' Data última manutenção.: 22/07/2004                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSql As String
Dim CNconexao As New DLLConexao_Sistema.Conexao
Dim CNconexao2 As New DLLConexao_Sistema.Conexao
Public strUsuario As String
Public strEstacao As String
Dim log As New DLLSystemManager.log

Private Sub cmdOk_Click()

    Dim rstLog As New ADODB.Recordset
    
    pgbAndamento.Visible = True
    
    'Conexão Log Morto
    CNconexao.Initial_Catalog = "BDLog_Morto"
    CNconexao.Abrir_conexao ("Otica")
    
    strSql = Empty
    strSql = "SELECT * FROM TBLog"
    Movimentacoes.Select_geral strSql, "BDLog", rstLog, "Otica", Me
        
    'Conexão Log
    CNconexao2.Initial_Catalog = "BDLog"
    CNconexao2.Abrir_conexao ("Otica")
    
    rstLog.MoveFirst
    
    If rstLog.BOF = True And rstLog.EOF = True Then
       MsgBox "Não existem registros a serem limpos no banco de LOG", vbInformation, "Logicx"
       Set rstLog = Nothing
       CNconexao.Fechar_conexao
       CNconexao2.Fechar_conexao
       Exit Sub
    End If
    
    'Progress Bar
    pgbAndamento.Min = 0
    pgbAndamento.Max = rstLog.RecordCount
    
    CNconexao.CNconexao.BeginTrans
    CNconexao2.CNconexao.BeginTrans
    
    On Error GoTo Erro
    
    Do While Not rstLog.EOF
    
        'LOG MORTO
        strSql = Empty
        strSql = "INSERT INTO TBLog_Morto(IXData_TBLog,DFTipo_TBLog,IXUsuario_TBLog,DFDescricao_TBLog,DFEstacao_TBLog,DFPrograma_TBLog,DFEvento_TBlog,DFHora_TBLog)" & _
                 "VALUES('" & Format(rstLog!IXData_TBLog, "YYYYMMDD") & "'," & rstLog!DFTipo_TBLog & ",'" & rstLog!IXUsuario_TBLog & "','" & rstLog!DFDescricao_TBLog & "','" & rstLog!DFEstacao_TBLog & "','" & rstLog!DFPrograma_TBLog & "','" & rstLog!DFEvento_TBlog & "','" & rstLog!DFHora_TBLog & "')"
        'Incluindo Registro no LOG MORTO
        CNconexao.CNconexao.Execute strSql
        
        'LOG
        strSql = Empty
        strSql = "DELETE FROM TBLog WHERE PKId_TBLog = " & rstLog!PKId_TBLog & " "
        'Excluindo Registro no LOG
        CNconexao2.CNconexao.Execute strSql
        
        'LOG_ERRO
        strSql = Empty
        strSql = "DELETE FROM TBLog_erro"
        
        'Excluindo Registro no LOG_ERRO
        CNconexao2.CNconexao.Execute strSql
        
        pgbAndamento.Value = rstLog.AbsolutePosition
        pgbAndamento.Refresh
        
        rstLog.MoveNext
    Loop
    
    'Atualizando parametros log
    
    rstLog.MoveFirst
    
    Dim strProxima_Limpeza As String
    Dim rstParametros As New ADODB.Recordset
    
    strSql = "SELECT * FROM TBParametros_Log"
    Movimentacoes.Select_geral strSql, "BDLog", rstParametros, "Otica", Me
    strProxima_Limpeza = DateAdd("d", CInt(Trim(rstParametros!DFNumero_Dias_Proxima_Limpeza)), Now)
    
    strSql = Empty
    strSql = "UPDATE TBParametros_Log SET DFProxima_Limpeza_Log = '" & Format(CDate(strProxima_Limpeza), "YYYYMMDD") & "',DFUltima_Limpeza_Log = '" & Format(Now, "YYYYMMDD") & "',DFConfere_execucao = 'N' WHERE PKCodigo_TBInformacoes_base = 1"
    CNconexao2.CNconexao.Execute strSql
    
    'Informações para gravar o LOG
    'Informações Constantes para o log
    log.Usuario = strUsuario
    log.Programa = "Login do Sistema"
    log.Estacao = strEstacao
    
    'Informações Variaveis para o log
    log.Evento = "Login do Sistema"
    log.Descricao = "Rotina de Limpeza do Banco de Log - Total de " & rstLog.RecordCount & " reg(s)."
    log.Tipo = 2
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
        
    CNconexao2.CNconexao.CommitTrans
    CNconexao2.Fechar_conexao
        
    CNconexao.CNconexao.CommitTrans
    CNconexao.Fechar_conexao
    
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    Call Conta_Registros_Log
    
    MsgBox "Operação realizada com sucesso!Favor Reinicie a aplicação", vbInformation, "Logicx"
     
    Set rstLog = Nothing
    Set rstParametros = Nothing
    
    Unload Me
    
    Exit Sub
    
Erro:

    'ROLLBACK'S
    CNconexao.CNconexao.RollbackTrans
    CNconexao2.CNconexao.RollbackTrans
    
    MsgBox "Ocorreu um erro : " & Err.Description & ", no reg. " & rstLog!PKId_TBLog & "", vbCritical, "Logicx"
    
    End
    
    Exit Sub
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Conta_Registros_Log
End Sub

Private Function Conta_Registros_Log()

    Dim rstLog As New ADODB.Recordset
    Dim rstLog_Morto As New ADODB.Recordset
    
    strSql = Empty
    strSql = "SELECT count(*) as TOTAL FROM TBLog"
    Movimentacoes.Select_geral strSql, "BDLog", rstLog, "Otica", Me
    Me.lblTotal_Log.Caption = rstLog!TOTAL & " Registro(s)"
    
    strSql = Empty
    strSql = "SELECT count(*) as TOTAL FROM TBLog_Morto"
    Movimentacoes.Select_geral strSql, "BDLog_Morto", rstLog_Morto, "Otica", Me
    lblLog_Morto.Caption = rstLog_Morto!TOTAL & " Registro(s)"
    
    Set rstLog = Nothing
    Set rstLog_Morto = Nothing

End Function

Private Sub Form_Unload(Cancel As Integer)
    strUsuario = Empty
    strEstacao = Empty
End Sub
