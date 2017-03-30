VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8b.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTeste_Aguarde 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   ScaleHeight     =   930
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pgbProgresso 
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   570
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash imgAguarde 
      Height          =   405
      Left            =   150
      TabIndex        =   1
      Top             =   90
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
      AllowScriptAccess=   ""
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
   Begin VB.Label lblArgumentos 
      Alignment       =   2  'Center
      Caption         =   "com argumentos"
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   570
      Width           =   2835
   End
End
Attribute VB_Name = "frmTeste_Aguarde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Nome_IMG = "por_favor_aguarde.swf"

Private Sub Form_Load()

    ReDim arrCmd(0) As String
    Dim I As Integer

    imgAguarde.Movie = App.Path & "\" & Nome_IMG
    imgAguarde.Play
    
    arrCmd = Split(Command(), ",")
    
    If Command() <> Empty Then
        If arrCmd(LBound(arrCmd)) <> Empty Then
            lblArgumentos.Visible = True
            lblArgumentos.Caption = lblArgumentos.Caption & " ("
            For I = LBound(arrCmd) To UBound(arrCmd)
                lblArgumentos.Caption = lblArgumentos.Caption & arrCmd(I) & ", "
                Select Case I
                    Case 0: pgbProgresso.Min = Trim(arrCmd(I))
                    Case 1: pgbProgresso.Max = Trim(arrCmd(I))
                    Case 2: pgbProgresso.Value = Trim(arrCmd(I))
                End Select
            Next I
            lblArgumentos.Caption = Mid(lblArgumentos.Caption, 1, InStrRev(lblArgumentos.Caption, ",") - 1) & ")"
        End If
    Else
        lblArgumentos.Visible = False
        lblArgumentos.Caption = Empty
        pgbProgresso.Min = 0
        pgbProgresso.Max = 1
        pgbProgresso.Value = 0
        pgbProgresso.Visible = False
        Me.Height = Me.Height - 330
    End If
    
    If Me.MDIChild = False Then
        Me.Left = Screen.Width - Me.Width - 300
        Me.Top = Screen.Height - Me.Height - 800
    End If
    
End Sub
