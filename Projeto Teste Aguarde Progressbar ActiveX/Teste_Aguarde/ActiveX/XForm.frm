VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8b.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form XForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   2835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pgbProgresso 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   690
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash imgAguarde 
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2835
      _cx             =   5001
      _cy             =   714
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Label lblAndamento 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   2835
   End
End
Attribute VB_Name = "XForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const Altura_Padrao_Form = 945
Const Nome_IMG = "por_favor_aguarde.swf"

Private Sub Form_Load()

    imgAguarde.Movie = App.Path & "\" & Nome_IMG
    imgAguarde.Play
    
'    xClass.Exibir_Barra_Progresso
'    Me.Height = Altura_Padrao_Form

    If Me.MDIChild = False Then
        
'        If Not Me.pgbProgresso.Visible Then
'            Me.Height = Me.Height - Me.lblAndamento.Height - Me.pgbProgresso.Height - 30
'        End If
        
        Me.Left = Screen.Width - Me.Width - 300
        Me.Top = Screen.Height - Me.Height - 800

    End If
    
    Me.lblAndamento.Caption = Me.pgbProgresso.Value & " / " & Me.pgbProgresso.Max
    
End Sub

'Public Sub setAltura_Padrao_Form()
'    Me.Height = Altura_Padrao_Form
'End Sub
