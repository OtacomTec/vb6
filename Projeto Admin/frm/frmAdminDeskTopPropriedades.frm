VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdminDesktopPropriedades 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gerenciador de Tarefas"
   ClientHeight    =   4815
   ClientLeft      =   4875
   ClientTop       =   3360
   ClientWidth     =   4110
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "frmAdminDeskTopPropriedades.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Line1(0)"
      Tab(0).Control(1)=   "Line1(1)"
      Tab(0).Control(2)=   "Image1"
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(4)=   "Label2(0)"
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(6)=   "lblID_AreaDeTrabalho"
      Tab(0).Control(7)=   "Label2(1)"
      Tab(0).Control(8)=   "Label2(2)"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Aplicativos"
      TabPicture(1)   =   "frmAdminDeskTopPropriedades.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lvwTaskMan"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmd(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmd(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmd 
         Caption         =   "&Finalizar Tarefa"
         Height          =   315
         Index           =   1
         Left            =   2775
         TabIndex        =   9
         Top             =   4350
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Alternar Para"
         Height          =   315
         Index           =   0
         Left            =   1530
         TabIndex        =   8
         Top             =   4350
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwTaskMan 
         Height          =   3825
         Left            =   90
         TabIndex        =   1
         Top             =   420
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   6747
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageListGeralPequeno"
         SmallIcons      =   "ImageListGeralPequeno"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tarefa"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageListGeralPequeno 
         Left            =   3480
         Top             =   -180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdminDeskTopPropriedades.frx":0038
               Key             =   "ico_Aplicativo"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         Height          =   195
         Index           =   2
         Left            =   -74820
         TabIndex        =   7
         Top             =   1950
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         Height          =   195
         Index           =   1
         Left            =   -74820
         TabIndex        =   6
         Top             =   1650
         Width           =   585
      End
      Begin VB.Label lblID_AreaDeTrabalho 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   -74010
         TabIndex        =   5
         Top             =   780
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         Height          =   195
         Left            =   -74250
         TabIndex        =   4
         Top             =   780
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   3
         Top             =   1350
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Área de Trabalho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74250
         TabIndex        =   2
         Top             =   480
         Width           =   1845
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -74850
         Picture         =   "frmAdminDeskTopPropriedades.frx":048A
         Top             =   480
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   -74880
         X2              =   -69316
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   -74865
         X2              =   -69316
         Y1              =   1065
         Y2              =   1065
      End
   End
End
Attribute VB_Name = "frmAdminDesktopPropriedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            'Unload Me
            AT.AlternarPara CLng(lvwTaskMan.SelectedItem.ListSubItems(2))
            
        Case 1
            AT.FinalizarTarefa CLng(lvwTaskMan.SelectedItem.ListSubItems(1))
            
    End Select
End Sub

Private Sub Form_Activate()
    'lvwTaskMan.c
    'frmAdminMDI.ActiveForm
End Sub

Private Sub ListarProgramas()
    lvwTaskMan.ColumnHeaders.Clear
    lvwTaskMan.ListItems.Clear
    lvwTaskMan.View = lvwReport
    ' Adicionando novas colunas
    lvwTaskMan.ColumnHeaders.Add , "Aplicativo", "Aplicativo", 2100
    lvwTaskMan.ColumnHeaders.Add , "PID", "PID", 850
    lvwTaskMan.ColumnHeaders.Add , "ID", "ID Janela", 850
    If frmAdminMDI.ActiveForm Is Nothing Then Exit Sub
    Dim itmX As ListItem
    Dim i As Integer
    For i = 1 To frmAdminMDI.ActiveForm.TotaldeProgramas
        Set itmX = lvwTaskMan.ListItems.Add(i, "Aplicativo " & i, frmAdminMDI.ActiveForm.Programa(0, i), "ico_Aplicativo", "ico_Aplicativo")
        itmX.SubItems(lvwTaskMan.ColumnHeaders("ID").SubItemIndex) = frmAdminMDI.ActiveForm.Programa(1, i)
        itmX.SubItems(lvwTaskMan.ColumnHeaders("PID").SubItemIndex) = frmAdminMDI.ActiveForm.Programa(2, i)
    Next i
End Sub

Private Sub Form_Load()
    'lvwTaskMan.imageicons = ImageListGeralPequeno
    ListarProgramas
End Sub

Private Sub lvwTaskMan_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'MsgBox Item.ListSubItems(1)
    
End Sub

