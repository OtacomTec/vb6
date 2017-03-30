VERSION 5.00
Begin VB.Form frmInterfaceSiac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programa de teste para a InterfaceSiac.dll"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmInterfaceSiac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDump 
      AutoRedraw      =   -1  'True
      Height          =   3855
      Left            =   240
      ScaleHeight     =   3795
      ScaleWidth      =   5835
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2280
      Width           =   5895
   End
   Begin VB.CommandButton cmdEncerraRecebimento 
      Caption         =   "Encerra Recebimento"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtPorta 
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "3000"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdIniciaRecebimento 
      Caption         =   "Inicia Recebimento"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdIniciaComunicacao 
      Caption         =   "Inicia Comunicação"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Timer tmrRecebe 
      Enabled         =   0   'False
      Left            =   3360
      Top             =   0
   End
   Begin VB.Label lblPorta 
      AutoSize        =   -1  'True
      Caption         =   "Porta: "
      Height          =   195
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   465
   End
End
Attribute VB_Name = "frmInterfaceSiac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' **********************************************************************
'                           P R O T Ó T I P O S
' **********************************************************************

' Função para iniciar a comunicação
Private Declare Function ComInicia Lib "InterfaceSiac.dll" (ByVal Porta As String) As Long

' Função para recebimento de mensagem
Private Declare Function ComRecebe Lib "InterfaceSiac.dll" ( _
    ByRef ComId As Integer, ByVal Transacao As String, ByRef Funcao As Integer, _
    ByRef RecBuffer As Any, ByRef RecLen As Integer, ByVal MaxBuf As Integer, _
    ByVal TimeOut As Long) As Long

' Função para envio de mensagem
Private Declare Function ComEnvia Lib "InterfaceSiac.dll" ( _
    ByVal ComId As Integer, ByRef EnvBuffer As Any, ByVal EnvLen As Integer) As Long

' **********************************************************************
'                               T I P O S
' **********************************************************************

' Tipo do buffer de recebimento da mensagem
Private Type TIPO_PEDIDO
    CodLoja As String * 4
    NumPdv As String * 3
    DataLocal As String * 6  ' (ddmmaa)
    HoraLocal As String * 6  ' (hhmmss)
    DataMovto As String * 6  ' (ddmmaa)
    NumSeqOp As String * 6
    TipoDoc As String * 1 ' 'C'-Cartão cliente   'D'-CPF/CGC
    CodCliente As String * 30 ' Delimitado por \0
End Type

' Tipo do buffer de envio da mensagem resposta - string para permitir cópia
Private Type TIPO_MSG_RESPOSTA
    Msg As String * 1997
End Type

' Tipo do buffer de envio da mensagem
Private Type TIPO_RESPOSTA
    CodResposta As String * 2   '00-OK     <> 00 - Erro
    IndRespGenerica As String * 1
    NumRespostas As String * 2
    Texto As TIPO_MSG_RESPOSTA
End Type

' **********************************************************************
'                           V A R I Á V E I S
' **********************************************************************

' Resposta dos pedidos
Dim Resposta As Long

' Dados para comunicação
Dim ComId As Integer
Dim Transacao As String * 4
Dim Funcao As Integer
Dim BufferPedido As TIPO_PEDIDO
Dim BufferResposta As TIPO_RESPOSTA
Dim TamMsg As Integer

' Contador das mensagens recebidas
Dim QtMsg As Long

' Nome do arquivo de log
Dim NomeArqLog As String

Private Sub cmdEncerraRecebimento_Click()
    ' Liga o timer de recebimento
    tmrRecebe.Enabled = False
End Sub

Private Sub cmdIniciaComunicacao_Click()
Dim strPorta As String * 6

    On Error GoTo Erro
    ' Monta a porta
    strPorta = Trim$(txtPorta.Text) + Chr(0)
    
    ' Inicia a comunicação
    Resposta = ComInicia(strPorta)
    
    ' Deu erro?
    If Resposta <> 0 Then
        MsgBox "A inicialização da comunicação não pôde ser efetuada!" & Chr(13) & "Erro: " & CStr(Resposta), vbOKOnly + vbExclamation, "InterfaceSiac"
    Else
        MsgBox "A inicialização foi efetuada com sucesso!", vbOKOnly + vbInformation, "InterfaceSiac"
    End If
    Exit Sub
    
Erro:
    MsgBox "Ocorreu o erro <" & CStr(Err.Number) & "> - " & Err.Description & Chr(13) & "no módulo " & Err.Source, vbOKOnly + vbExclamation, "InterfaceSiac"
    Resume ErroInterno
ErroInterno:
    MsgBox "A inicialização da comunicação não pôde ser efetuada!", vbOKOnly + vbExclamation, "InterfaceSiac"
End Sub

Private Sub cmdIniciaRecebimento_Click()
    ' Liga o timer de recebimento
    tmrRecebe.Enabled = True
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    ' Inicializa o tempo para o timer de recebimento
    ' 1000 milisegundos
    tmrRecebe.Interval = 1000
    QtMsg = 0
    ' Monta o nome do arquivo de log
    NomeArqLog = App.Path & "\" & App.EXEName & ".log"
    Kill NomeArqLog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If MsgBox("Deseja mesmo sair?", vbYesNo + vbDefaultButton2 + vbQuestion, "InterfaceSiac") _
        <> vbYes Then
        Cancel = True
    End If
End Sub

Private Sub tmrRecebe_Timer()
    On Error Resume Next
    ' Desliga o timer
    tmrRecebe.Enabled = False
    ' Verifica se chegou pedido
    VerificaPedido
End Sub

Private Sub VerificaPedido()
    
    On Error GoTo Erro
    ' Chegou alguma mensagem?
    GravaLog "Vai tentar receber mensagem..."
    Resposta = ComRecebe(ComId, Transacao, Funcao, BufferPedido, TamMsg, _
        Len(BufferPedido), 200)
    If Resposta <> 0 Then
        GravaLog "    Não chegou mensagem. Resposta: " & CStr(Resposta)
        ' Não recebeu a mensagem
        picDump.Cls
        picDump.Print "ComRecebe respondeu com... " & CStr(Resposta)
        ' Liga o timer novamente
        tmrRecebe.Enabled = True
        Exit Sub
    End If
    
    GravaLog "    Chegou mensagem. Resposta: " & CStr(Resposta)
    
    If QtMsg > 9999 Then
        QtMsg = 0
    Else
        QtMsg = QtMsg + 1
    End If
    
    ' Dump
    picDump.Cls
    picDump.Print "Recebeu uma mensagem! É a de número " & CStr(QtMsg) & Chr(13) & "Retorno: " & CStr(Resposta)
    GravaLog "        " & "Recebeu uma mensagem! É a de número " & CStr(QtMsg) & Chr(13) & "Retorno: " & CStr(Resposta)
    picDump.Print
    GravaLog "        " & ""
    picDump.Print "Id: " & CStr(ComId) & " Transação: " & CStr(Transacao) & " Funcao: " & CStr(Funcao) & " TamMsg: " & CStr(TamMsg)
    GravaLog "        " & "Id: " & CStr(ComId) & " Transação: " & CStr(Transacao) & " Funcao: " & CStr(Funcao) & " TamMsg: " & CStr(TamMsg)
    picDump.Print
    GravaLog "        " & ""
    picDump.Print "Código da loja...................: " & BufferPedido.CodLoja
    GravaLog "        " & "Código da loja...................: " & BufferPedido.CodLoja
    picDump.Print "Número do Pdv....................: " & BufferPedido.NumPdv
    GravaLog "        " & "Número do Pdv....................: " & BufferPedido.NumPdv
    picDump.Print "Data local.......................: " & BufferPedido.DataLocal
    GravaLog "        " & "Data local.......................: " & BufferPedido.DataLocal
    picDump.Print "Hora local.......................: " & BufferPedido.HoraLocal
    GravaLog "        " & "Hora local.......................: " & BufferPedido.HoraLocal
    picDump.Print "Data do movimento................: " & BufferPedido.DataMovto
    GravaLog "        " & "Data do movimento................: " & BufferPedido.DataMovto
    picDump.Print "Número sequencial de operação....: " & BufferPedido.NumSeqOp
    GravaLog "        " & "Número sequencial de operação....: " & BufferPedido.NumSeqOp
    picDump.Print "Tipo do documento................: " & BufferPedido.TipoDoc
    GravaLog "        " & "Tipo do documento................: " & BufferPedido.TipoDoc
    picDump.Print "Código do cliente................: " & BufferPedido.CodCliente
    GravaLog "        " & "Código do cliente................: " & BufferPedido.CodCliente
 
    ' Montar a resposta
    GravaLog ""
    BufferResposta.CodResposta = "00"  ' OK
    BufferResposta.IndRespGenerica = "0"
    BufferResposta.NumRespostas = "05"
    BufferResposta.Texto.Msg = "01DA033Mensagem com display na abertura." & _
        "01IF050Mensagem impressa no final do cupom no fechamento." & _
        "03MA032Mensagem no monitor na abertura." & _
        "04T 059Mensagem a ser mostrada no monitor tela cheia.@Pulou linha." & _
        "05C 0044321"
        
    ' Envia a resposta
    GravaLog "        " & "Vai chamar ComEnvia"
    Resposta = ComEnvia(ComId, BufferResposta, Len(BufferResposta.CodResposta) + _
        Len(BufferResposta.IndRespGenerica) + Len(BufferResposta.NumRespostas) + _
        Len(Trim$(BufferResposta.Texto.Msg)))
    GravaLog "        " & "Chamou ComEnvia"
    picDump.Print
    picDump.Print "Resposta... " & CStr(Resposta)
    GravaLog "        " & "Resposta... " & CStr(Resposta)
    Exit Sub
Erro:
    MsgBox "Ocorreu o erro <" & CStr(Err.Number) & "> - " & Err.Description & Chr(13) & "no módulo " & Err.Source, vbOKOnly + vbExclamation, "InterfaceSiac"
    Resume ErroInterno
ErroInterno:
    MsgBox "A Verificação do pedido não foi efetuada!", vbOKOnly + vbExclamation, "InterfaceSiac"
End Sub

Private Sub GravaLog(strMsg As String)
Dim File As Integer
    On Error Resume Next
    File = FreeFile()
    Open NomeArqLog For Append Access Write As #File
    Print #File, Now & " - "; strMsg
    Close #File
End Sub
