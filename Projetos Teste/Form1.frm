VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   Caption         =   "Formulário Padrão"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox mskCPF 
      Height          =   330
      Left            =   1155
      TabIndex        =   5
      Top             =   1680
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   14
      Mask            =   "###.###.###-##"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controle"
      ForeColor       =   &H00800000&
      Height          =   1065
      Left            =   120
      TabIndex        =   22
      Top             =   4305
      Width           =   5265
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Height          =   330
         Left            =   210
         TabIndex        =   7
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   330
         Left            =   1890
         TabIndex        =   12
         Top             =   630
         Width           =   855
      End
      Begin VB.CommandButton cmdUltimo 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4200
         TabIndex        =   16
         Top             =   630
         Width           =   855
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "Alterar"
         Height          =   330
         Left            =   210
         TabIndex        =   10
         Top             =   630
         Width           =   855
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   330
         Left            =   1050
         TabIndex        =   11
         Top             =   630
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   330
         Left            =   1890
         TabIndex        =   9
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton cmdAnterior 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         TabIndex        =   13
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton cmdPrimeiro 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         TabIndex        =   15
         Top             =   630
         Width           =   855
      End
      Begin VB.CommandButton cmdPosterior 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4200
         TabIndex        =   14
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar"
         Height          =   330
         Left            =   1050
         TabIndex        =   8
         Top             =   315
         Width           =   855
      End
   End
   Begin MSComCtl2.DTPicker dtpData_ult_entrada 
      Height          =   330
      Left            =   3360
      TabIndex        =   6
      Top             =   1680
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   582
      _Version        =   393216
      Format          =   22675457
      CurrentDate     =   37608
   End
   Begin MSDataListLib.DataCombo dtcFornecedor 
      Height          =   360
      Left            =   1050
      TabIndex        =   3
      Top             =   945
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483644
      ForeColor       =   12582912
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dtgProduto 
      Height          =   1800
      Left            =   120
      TabIndex        =   20
      Top             =   2415
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   3175
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtEstoque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtcod_Fornecedor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   945
      Width           =   855
   End
   Begin VB.TextBox txtDescricao_produto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1035
      TabIndex        =   1
      Top             =   315
      Width           =   4320
   End
   Begin VB.TextBox txtcod_produto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   315
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "CPF do Produto"
      Height          =   195
      Left            =   1155
      TabIndex        =   24
      Top             =   1470
      Width           =   1125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Listagem de Produtos"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   2205
      Width           =   1530
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Data da Ùltima Entrada"
      Height          =   195
      Left            =   3360
      TabIndex        =   21
      Top             =   1470
      Width           =   1650
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Estoque"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   1470
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   735
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Produto"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   105
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   *** Formulário de teste onde usa-se todos os módulos                                   '
'                                                                                          '
'                                                                                          '
'                                                                                          '
'                                                                                          '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Sub cmdConfirmar_Click()
  Dim campos As String
  Dim valores As String
  
  campos = "PKCod_Produto,DFDescricao_Produto,FKCodigo_Fornecedor,DFEstoque,DFData_Ultima_Entrada,DFCPF_Produto"
  valores = " " & txtcod_produto.Text & "," & txtDescricao_produto.Text & "," & Me.txtcod_Fornecedor.Text & "," & Me.txtEstoque & "," & Me.dtpData_ult_entrada.Value & "," & mskCPF.Text & " "
  
  Call Botoes.Confirmar(Me, "TBProduto", campos, "Formulário Padrão", txtcod_produto, valores, "N")
  
End Sub

Private Sub cmdIncluir_Click()
  Me.txtcod_Fornecedor = Empty
End Sub

Private Sub dtcFornecedor_Change()
   txtcod_Fornecedor.Text = dtcFornecedor.BoundText
End Sub

Private Sub Form_Load()

  Dim strSql As String
  
  Call Conexao_Banco.Abre_Conexao("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Projetos\Projeto Teste\BDTeste.mdb;Persist Security Info=False")
  Call Movimentacoes.Inicio("PKCod_Fornec", "DFDescricao", "TBFornecedor", dtcFornecedor)
  strSql = "SELECT PKCod_Produto,DFDescricao_Produto FROM TBProduto"
  Call Movimentacoes.Inicio_DataGrid(strSql, dtgProduto, "1500,5000", "Código,Descrição")
    
  
  
  
  
  
End Sub

Private Sub txtcod_Fornecedor_LostFocus()
  dtcFornecedor.BoundText = txtcod_Fornecedor.Text
End Sub
