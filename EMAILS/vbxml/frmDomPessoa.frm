VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDomPessoa 
   Caption         =   "Usando XMLDOM com VB"
   ClientHeight    =   3810
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdxml 
      Caption         =   "XML >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   22
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   3360
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Caption         =   "Nós Filhos  (Elements) / Pessoas"
      Height          =   2295
      Left            =   3240
      TabIndex        =   9
      Top             =   720
      Width           =   3735
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1140
         TabIndex        =   14
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   1140
         TabIndex        =   13
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtTel 
         Height          =   285
         Left            =   1140
         TabIndex        =   12
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtEndereco 
         Height          =   285
         Left            =   1140
         TabIndex        =   11
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   1140
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblChildTag 
         Caption         =   "Email"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblChildTag 
         Caption         =   "Fax"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblChildTag 
         Caption         =   "Tel"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblChildTag 
         Caption         =   "Endereço :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblChildTag 
         Caption         =   "Nome"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Top             =   0
      Width           =   3735
      Begin VB.Label lblElemento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Elemento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Incluir/Salvar/Excluir NOVO Item"
      Height          =   675
      Left            =   3240
      TabIndex        =   3
      Top             =   3120
      Width           =   3795
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   315
         Left            =   2520
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "&Salvar "
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "&Incluir "
         Enabled         =   0   'False
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin SHDocVwCtl.WebBrowser webTarget 
      Height          =   2835
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   6615
      ExtentX         =   11668
      ExtentY         =   5001
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
      Location        =   "res://C:\WINNT\system32\shdoclc.dll/dnserror.htm#http:///"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView tvwPessoa 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5530
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdPreencher 
      Caption         =   "Preenche"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6720
      Y1              =   3840
      Y2              =   3840
   End
End
Attribute VB_Name = "frmDomPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_objDOMPessoa As DOMDocument
Private m_blnItemClicked As Boolean
Private m_strXmlPath As String
Dim flagpreenche As Boolean
Private Sub PreencherTreeWithChildren(objDOMNode As IXMLDOMElement)

  Dim objNameNode As IXMLDOMNode
  Dim objAttributes As IXMLDOMNamedNodeMap
  Dim objAttributeNode As IXMLDOMNode
  Dim objPessoaElement As IXMLDOMElement
  Dim intIndex As Integer
  
  Dim tvwElement As Node
  Dim tvwChildElement As Node

  'obtem o nome do elemento selecionado
  Set objNameNode = objDOMNode.selectSingleNode("NAME")

  'inclui os elementos aos nós
  Set tvwElement = tvwPessoa.Nodes.Add(1, tvwChild)
  tvwElement.Text = objNameNode.parentNode.nodeName & ": " _
                    & objNameNode.nodeTypedValue
  
  Set objAttributes = objDOMNode.Attributes
  
  'verifica os atributos
  If objAttributes.length > 0 Then
   
    ' obtendo o item para a referencia  ''PERSONID',
    ' com NameNodeListMap para o Nó atual
    Set objAttributeNode = objAttributes.getNamedItem("PERSONID")
    
    'armazena o valor na tag do treeview
    tvwElement.Tag = objAttributeNode.nodeValue
  End If
  
  tvwElement.EnsureVisible
  intIndex = tvwElement.Index
  
  'interagem através dos Nós filhos(childNodes) do objeto DOMNode
  ' para preencher o TreeView os seus valores
  For Each objPessoaElement In objDOMNode.childNodes
    Set tvwChildElement = tvwPessoa.Nodes.Add(intIndex, tvwChild)
    tvwChildElement.Text = objPessoaElement.nodeTypedValue
  Next
       
End Sub

Private Sub cmdExcluir_Click()
  deleteSelectedPerson tvwPessoa.SelectedItem
End Sub

Private Sub cmdIncluir_Click()
  lblElemento.Caption = ""
  txtNome.Text = ""
  txtEndereco.Text = ""
  txtTel.Text = ""
  txtFax.Text = ""
  txtEmail.Text = ""
  cmdIncluir.Enabled = True

End Sub

Private Sub cmdPreencher_Click()
  Dim objPessoaRoot As IXMLDOMElement
  Dim objPessoaElement As IXMLDOMElement
  Dim tvwRoot As Node
  Dim X As IXMLDOMNodeList
  
  flagpreenche = True
  cmdxml.Enabled = True
    
  Set m_objDOMPessoa = New DOMDocument
  
  m_objDOMPessoa.resolveExternals = True
  
  m_objDOMPessoa.validateOnParse = True
      
  'carrega o XML no documento DOM
  m_objDOMPessoa.async = False
  Call m_objDOMPessoa.Load(m_strXmlPath)
  
  'verifica se a carga do XML foi feita com sucesso
  If m_objDOMPessoa.parseError.reason <> "" Then
    MsgBox m_objDOMPessoa.parseError.reason
    Exit Sub
  End If
  
  'obtem o elemento raiz do XML
   Set objPessoaRoot = m_objDOMPessoa.documentElement
  
 
  'define as propriedades do Treeview
  tvwPessoa.LineStyle = tvwRootLines
  tvwPessoa.Style = tvwTreelinesPlusMinusText
  tvwPessoa.Indentation = 400
  
  'verifica se o treeview ja foi preenchido
  'se ja foi remove o raiz que remove tudo
  If tvwPessoa.Nodes.Count > 0 Then
    tvwPessoa.Nodes.Remove 1
  End If
  
  ' inclui um no filho ao no rai do TreeView
  Set tvwRoot = tvwPessoa.Nodes.Add()
  tvwRoot.Text = objPessoaRoot.baseName
  
  ' iteracao atraves de cada elemento para encher a arvore
  ' que por sua vez interaagem atraves de cada childNode
  ' do element(objPessoaElement)
  For Each objPessoaElement In objPessoaRoot.childNodes
    PreencherTreeWithChildren objPessoaElement
  Next
  
  webTarget.Navigate m_strXmlPath
  cmdExcluir.Enabled = True
  cmdIncluir.Enabled = True

End Sub

Private Sub cmdSair_Click()
   If MsgBox("Quer mesmo sair ?", vbYesNo, "Saida") = vbYes Then
       End
   End If
End Sub

Private Sub cmdSalvar_Click()
    salvarNovaPessoa
    cmdAdd.Enabled = False
End Sub

Private Sub Cmdxml_Click()
If cmdxml.Caption = "XML >>" Then
  Me.Height = 7200
  webTarget.Visible = True
  cmdxml.Caption = "<< XML"
ElseIf cmdxml.Caption = "<< XML" Then
  Me.Height = 4215
  webTarget.Visible = False
  cmdxml.Caption = "XML >>"
End If
End Sub

Private Sub Form_Load()
  m_strXmlPath = App.Path & "\agenda.xml"
  flagpreenche = False
End Sub

Private Sub tvwPessoa_Click()
  Dim objSelNode As Node
  
  If m_blnItemClicked = True Then
    m_blnItemClicked = False
    Exit Sub
  End If
  
  Set objSelNode = tvwPessoa.SelectedItem
  populatePessoaDetalhes objSelNode
    
End Sub

Private Sub populatePessoaDetalhes(objSelNode As Node)

' preenche os textbox dos formulario com os detalhes do documento
  Dim objPessoaElement As IXMLDOMElement
  Dim objChildElement As IXMLDOMElement
  
  If objSelNode Is Nothing Then Exit Sub
  
  'ignora a selecao do TreeView  se nao foi clicado
  'um No "PERSON"
  If Trim(objSelNode.Tag) <> "" Then
  
    'obtem o no(element type), que possui um atributo ao valor
    'da tag do TreeView
    Set objPessoaElement = m_objDOMPessoa.nodeFromID(objSelNode.Tag)
    
    lblElemento.Caption = objPessoaElement.nodeName & ": ID = " & _
                         objPessoaElement.Attributes(0).nodeValue
                         
    'interagem atraves dos nos achadose preenche
    'o texto com os conteudo dos nos
    For Each objChildElement In objPessoaElement.childNodes
    
      'verifica o tipo de No que estamos tratando
      If objChildElement.nodeType = NODE_ELEMENT Then
        Select Case UCase(objChildElement.nodeName)
        Case "NAME"
            txtNome.Text = objChildElement.nodeTypedValue
        Case "ADDRESS"
            txtEndereco.Text = objChildElement.nodeTypedValue
        Case "TEL"
            txtTel.Text = objChildElement.nodeTypedValue
        Case "FAX"
            txtFax.Text = objChildElement.nodeTypedValue
        Case "EMAIL"
            txtEmail.Text = objChildElement.nodeTypedValue
        End Select
      End If
    Next objChildElement
  End If

   Set objChildElement = Nothing
  Set objPessoaElement = Nothing
End Sub

Private Sub salvarNovaPessoa()

'cria um novo elemento e inclui no objeto dom
  Dim objPerson As IXMLDOMElement
  Dim objNewChild As IXMLDOMElement
  
  'cria um novo elemento PERSON
  Set objPerson = m_objDOMPessoa.createElement("PERSON")
  objPerson.setAttribute "PERSONID", getNewID
  m_objDOMPessoa.documentElement.appendChild objPerson
  
  'cria um elmeneto (objPerson), e inclui no No Filho(childNodes)
  Set objNewChild = m_objDOMPessoa.createElement("NAME")
  objNewChild.Text = txtName.Text
  objPerson.appendChild objNewChild
  
  Set objNewChild = m_objDOMPessoa.createElement("ADDRESS")
  objNewChild.Text = txtAddress.Text
  objPerson.appendChild objNewChild
  
  Set objNewChild = m_objDOMPessoa.createElement("TEL")
  objNewChild.Text = txtTel.Text
  objPerson.appendChild objNewChild
  
  Set objNewChild = m_objDOMPessoa.createElement("FAX")
  objNewChild.Text = txtFax.Text
  objPerson.appendChild objNewChild
  
  Set objNewChild = m_objDOMPessoa.createElement("EMAIL")
  objNewChild.Text = txtEmail.Text
  objPerson.appendChild objNewChild
  
  'sincroniza com o TreeView
  PreencherTreeWithChildren objPerson
  
  m_objDOMPessoa.save m_strXmlPath
  
  webTarget.Refresh
  Set objPerson = Nothing
  Set objNewChild = Nothing

End Sub
Private Function getNewID() As String
  getNewID = "p" & m_objDOMPessoa.documentElement.childNodes.length + 1
End Function

Private Sub tvwPessoa_Collapse(ByVal Node As MSComctlLib.Node)
  populatePessoaDetalhes Node
  m_blnItemClicked = True
End Sub

Private Sub tvwPessoa_Expand(ByVal Node As MSComctlLib.Node)
  populatePessoaDetalhes Node
  m_blnItemClicked = True
End Sub

Private Sub deleteSelectedPerson(objSelNode As Node)

  Dim objPessoaElement As IXMLDOMNode
  Dim objChildElement As IXMLDOMElement
  
  'se não selecionou um no na arvore sai
  If objSelNode Is Nothing Then Exit Sub
  
  'acha o no atual no TreeView
  'ou o seu Pai que possui um valor atribuido a tag
  If Trim(objSelNode.Tag) = "" Then
    If Trim(objSelNode.Parent.Tag) <> "" Then
      Set objSelNode = objSelNode.Parent
    End If
  End If
  
  If Trim(objSelNode.Tag) <> "" Then
  
    Set objPessoaElement = m_objDOMPessoa.nodeFromID(objSelNode.Tag)
    
    'remove o no do DOMDocument encontrado
    m_objDOMPessoa.documentElement.removeChild objPessoaElement
    m_objDOMPessoa.save m_strXmlPath
    tvwPessoa.Nodes.Remove objSelNode.Index
    webTarget.Refresh
  End If
    
End Sub

