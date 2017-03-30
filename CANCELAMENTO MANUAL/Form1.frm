VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "CANCELAMENTO MANUAL"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wbStatus_remessa 
      Height          =   3945
      Left            =   90
      TabIndex        =   5
      Top             =   1080
      Width           =   6765
      ExtentX         =   11933
      ExtentY         =   6959
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox txtProtocolo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1410
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtChave_acesso 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1410
      TabIndex        =   1
      Top             =   90
      Width           =   3465
   End
   Begin VB.CommandButton cmdCancelamento 
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5040
      TabIndex        =   0
      Top             =   390
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Protocolo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   510
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Chave Acesso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function PROCESSA_NFE() As Boolean

    Dim Cancela As cancNFe
    Dim NFeHelper As NFeHelper
    Dim WSHelper As WSHelper
    Dim Cancelamento As String
    Dim strData As String
    Dim strHora As String
    Dim strValidade As String
    Dim strValor_chave_rec As String
    Dim NumArq As Integer
    Dim strsql As String
    Dim rstParametros_venda As New ADODB.Recordset
    
    strsql = Empty
    strsql = "SELECT DFBaixa_estoque_composicao_TBParametros_venda,DFCaminho_danf_TBParametro_venda,DFIdentificacao_certificado_TBParametro_venda FROM TBParametros_venda WHERE IXCodigo_TBEmpresa = 100"
    Movimentacoes.Select_geral strsql, "BDRetaguarda", rstParametros_venda, "Otica", Me
    
    strPath_DANF = rstParametros_venda!DFCaminho_danf_TBParametro_venda
    Certificate = Funcoes_Gerais.Abrir_NFE_registro("Otica", Me)
    
    wbStatus_remessa.Navigate Funcoes_Gerais.Abrir_figura_registro("Otica", Me) & "\PRE_XMLPROC.HTM"
    
    PROCESSA_NFE = True

    Set Cancela = New cancNFe

    Cancela.versao = "1.07"
    Cancela.infCanc.xServ = "CANCELAR"
    Cancela.infCanc.xJust = "CANCELAMENTO MANUAL POR MOTIVOS DIVERSOS"
    
'    If MDIPrincipal.booDesign_time = True Then
'       Cancela.infCanc.tpAmb = "2"
'    Else
       Cancela.infCanc.tpAmb = "1"
'    End If
    
    Cancela.infCanc.Id = "ID" + txtChave_acesso.Text
    Cancela.infCanc.chNFe = txtChave_acesso.Text
    Cancela.infCanc.nProt = txtProtocolo.Text
    
'    If MDIPrincipal.booDesign_time = True Then
'       booProducao = False
'    Else
'       booProducao = True
'    End If
    
    strData = Format(Now, "YYYYMMDD")
    strHora = Format(Now, "hhmm")
    
    Set NFeHelper = New NFeHelper
    Set WSHelper = New WSHelper

    Cancelamento = ""
    If (NFeHelper.GenerateXml_4(Cancela, Certificate)) Then

        Cancelamento = Cancela.XMLGenerated
        'txtXml.Text = Cancelamento
        strValidade = NFeHelper.ValidateXML(Cancelamento, Certificate)
        
        If (Len(strValidade) = 0) Then
            NumArq = FreeFile
            Open strPath_DANF & "\RET\" & strHora & strData & "-XML_RECEITA-CANC-PROCESS" & ".XML" For Append As #NumArq
            Print #NumArq, WSHelper.CallNfeCancelamentoNFWebMethod(Cancelamento, Certificate, WS_TARGET.WS_TARGET_VR, True)
            Close #NumArq
            PROCESSA_NFE = True
        Else
            NumArq = FreeFile
            Open strPath_DANF & "\RET\" & strHora & strData & "-XML_RECEITA-CANC-PROCESS" & ".XML" For Append As #NumArq
            Print #NumArq, "Comando não enviado. Erro no schema!"
            Close #NumArq
            PROCESSA_NFE = False
        End If
        
'        'LER XML E VERIFICAR SE FOI CANCELADO OK....SENÃO REJEITAR O CANCELAMENTO TODO.
'        strValor_chave_rec = LER_XML_PROC(strPath_DANF & "\RET\" & strHora & strData & "-XML_RECEITA-CANC-PROCESS" & ".XML")
'
        If strValor_chave_rec = "Cancelamento de NF-e homologado" Then
           wbStatus_remessa.Navigate strPath_DANF & "\RET\" & strHora & strData & "-XML_RECEITA-CANC-PROCESS" & ".XML"
           PROCESSA_NFE = True
        Else
           wbStatus_remessa.Navigate strPath_DANF & "\RET\" & strHora & strData & "-XML_RECEITA-CANC-PROCESS" & ".XML"
           PROCESSA_NFE = False
        End If
        
    Else
        NumArq = FreeFile
        Open strPath_DANF & "\RET\" & strHora & strData & "-XML_RECEITA-CANC-PROCESS" & ".XML" For Append As #NumArq
        Print #NumArq, "Comando não enviado. Erro no schema!"
        Close #NumArq
        PROCESSA_NFE = False
        wbStatus_remessa.Navigate strPath_DANF & "\RET\" & strHora & strData & "-XML_RECEITA-CANC-PROCESS" & ".XML"
    End If

End Function

Private Sub cmdCancelamento_Click()
    Call PROCESSA_NFE
End Sub

