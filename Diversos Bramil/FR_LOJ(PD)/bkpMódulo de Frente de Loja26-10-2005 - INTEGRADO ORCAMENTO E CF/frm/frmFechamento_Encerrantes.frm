VERSION 5.00
Begin VB.Form frmFechamento_Encerrantes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdProx_bombas 
      BackColor       =   &H0080FFFF&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8340
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   540
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8340
      Width           =   1455
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   17
      Left            =   8400
      TabIndex        =   16
      Top             =   2490
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   19
      Left            =   8400
      TabIndex        =   18
      Top             =   3390
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   18
      Left            =   10020
      TabIndex        =   17
      Top             =   2490
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   20
      Left            =   10020
      TabIndex        =   19
      Top             =   3390
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   24
      Left            =   10020
      TabIndex        =   23
      Top             =   6900
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   22
      Left            =   10020
      TabIndex        =   21
      Top             =   6000
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   23
      Left            =   8400
      TabIndex        =   22
      Top             =   6900
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   21
      Left            =   8400
      TabIndex        =   20
      Top             =   6000
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   9
      Left            =   4560
      TabIndex        =   8
      Top             =   2490
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   11
      Left            =   4560
      TabIndex        =   10
      Top             =   3390
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   10
      Left            =   6180
      TabIndex        =   9
      Top             =   2490
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   12
      Left            =   6180
      TabIndex        =   11
      Top             =   3390
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   16
      Left            =   6180
      TabIndex        =   15
      Top             =   6900
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   14
      Left            =   6180
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   15
      Left            =   4560
      TabIndex        =   14
      Top             =   6900
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   13
      Left            =   4560
      TabIndex        =   12
      Top             =   6000
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   750
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   750
      TabIndex        =   6
      Top             =   6900
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   2370
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   8
      Left            =   2370
      TabIndex        =   7
      Top             =   6900
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   2370
      TabIndex        =   3
      Top             =   3390
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   2370
      TabIndex        =   1
      Top             =   2490
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   750
      TabIndex        =   2
      Top             =   3390
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   750
      TabIndex        =   0
      Top             =   2490
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   11520
      Picture         =   "frmFechamento_Encerrantes.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   435
      TabIndex        =   27
      Top             =   30
      Width           =   435
   End
   Begin VB.Shape Shape38 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   495
      Left            =   9960
      Shape           =   4  'Rounded Rectangle
      Top             =   8400
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Shape Shape37 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   495
      Left            =   450
      Shape           =   4  'Rounded Rectangle
      Top             =   8400
      Width           =   1485
   End
   Begin VB.Label lblPDV 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "PDV"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   6915
      TabIndex        =   59
      Top             =   8430
      Width           =   585
   End
   Begin VB.Label lblOperador 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Operador:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   2250
      TabIndex        =   58
      Top             =   8430
      Width           =   1440
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   17
      Left            =   8310
      Shape           =   4  'Rounded Rectangle
      Top             =   2580
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   19
      Left            =   8310
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   18
      Left            =   9930
      Shape           =   4  'Rounded Rectangle
      Top             =   2580
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   20
      Left            =   9930
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico17"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   17
      Left            =   8400
      TabIndex        =   57
      Top             =   2220
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   20
      Left            =   10020
      TabIndex        =   56
      Top             =   3120
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico18"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   18
      Left            =   10020
      TabIndex        =   55
      Top             =   2220
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico19"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   19
      Left            =   8400
      TabIndex        =   54
      Top             =   3120
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBomba 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBomba5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   5
      Left            =   8400
      TabIndex        =   53
      Top             =   1710
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblBomba 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBomba6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   6
      Left            =   8400
      TabIndex        =   52
      Top             =   5220
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico23"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   23
      Left            =   8400
      TabIndex        =   51
      Top             =   6630
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico22"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   22
      Left            =   10020
      TabIndex        =   50
      Top             =   5730
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico24"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   24
      Left            =   10020
      TabIndex        =   49
      Top             =   6630
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico21"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   21
      Left            =   8400
      TabIndex        =   48
      Top             =   5730
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   24
      Left            =   9930
      Shape           =   4  'Rounded Rectangle
      Top             =   6990
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   22
      Left            =   9930
      Shape           =   4  'Rounded Rectangle
      Top             =   6090
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   23
      Left            =   8310
      Shape           =   4  'Rounded Rectangle
      Top             =   6990
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   21
      Left            =   8310
      Shape           =   4  'Rounded Rectangle
      Top             =   6090
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   9
      Left            =   4470
      Shape           =   4  'Rounded Rectangle
      Top             =   2580
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   11
      Left            =   4470
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   10
      Left            =   6090
      Shape           =   4  'Rounded Rectangle
      Top             =   2580
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   12
      Left            =   6090
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   9
      Left            =   4560
      TabIndex        =   47
      Top             =   2220
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   12
      Left            =   6180
      TabIndex        =   46
      Top             =   3120
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   10
      Left            =   6180
      TabIndex        =   45
      Top             =   2220
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   11
      Left            =   4560
      TabIndex        =   44
      Top             =   3120
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBomba 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBomba3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   4560
      TabIndex        =   43
      Top             =   1710
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblBomba 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBomba4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   4
      Left            =   4560
      TabIndex        =   42
      Top             =   5220
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   15
      Left            =   4560
      TabIndex        =   41
      Top             =   6630
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico14"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   14
      Left            =   6180
      TabIndex        =   40
      Top             =   5730
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico16"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   16
      Left            =   6180
      TabIndex        =   39
      Top             =   6630
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   13
      Left            =   4560
      TabIndex        =   38
      Top             =   5730
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   16
      Left            =   6090
      Shape           =   4  'Rounded Rectangle
      Top             =   6990
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   14
      Left            =   6090
      Shape           =   4  'Rounded Rectangle
      Top             =   6090
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   15
      Left            =   4470
      Shape           =   4  'Rounded Rectangle
      Top             =   6990
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   13
      Left            =   4470
      Shape           =   4  'Rounded Rectangle
      Top             =   6090
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   5
      Left            =   660
      Shape           =   4  'Rounded Rectangle
      Top             =   6090
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   7
      Left            =   660
      Shape           =   4  'Rounded Rectangle
      Top             =   6990
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   6
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   6090
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   8
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   6990
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   5
      Left            =   750
      TabIndex        =   37
      Top             =   5730
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   8
      Left            =   2370
      TabIndex        =   36
      Top             =   6630
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   6
      Left            =   2370
      TabIndex        =   35
      Top             =   5730
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   7
      Left            =   750
      TabIndex        =   34
      Top             =   6630
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblBomba 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBomba2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   750
      TabIndex        =   33
      Top             =   5220
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblBomba 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBomba1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   750
      TabIndex        =   32
      Top             =   1710
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   750
      TabIndex        =   31
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   2370
      TabIndex        =   30
      Top             =   2220
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   4
      Left            =   2370
      TabIndex        =   29
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblBico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "lblBico1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   750
      TabIndex        =   28
      Top             =   2220
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   4
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   2
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   2580
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   3
      Left            =   660
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bico 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   405
      Index           =   1
      Left            =   660
      Shape           =   4  'Rounded Rectangle
      Top             =   2580
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape Shape_bomba 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2925
      Index           =   1
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Shape Shape_bomba2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2985
      Index           =   1
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   1620
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fechamento Encerrantes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   1800
      TabIndex        =   26
      Top             =   540
      Width           =   6525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      X1              =   4950
      X2              =   0
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Shape Shape_bomba 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2925
      Index           =   2
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   5070
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Shape Shape_bomba2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2985
      Index           =   2
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   5130
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Shape Shape_bomba 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2925
      Index           =   3
      Left            =   4290
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Shape Shape_bomba2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2985
      Index           =   3
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   1620
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Shape Shape_bomba 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2925
      Index           =   4
      Left            =   4290
      Shape           =   4  'Rounded Rectangle
      Top             =   5070
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Shape Shape_bomba2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2985
      Index           =   4
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   5130
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Shape Shape_bomba 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2925
      Index           =   5
      Left            =   8130
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Shape Shape_bomba2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2985
      Index           =   5
      Left            =   8040
      Shape           =   4  'Rounded Rectangle
      Top             =   1620
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Shape Shape_bomba 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2925
      Index           =   6
      Left            =   8130
      Shape           =   4  'Rounded Rectangle
      Top             =   5070
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Shape Shape_bomba2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2985
      Index           =   6
      Left            =   8040
      Shape           =   4  'Rounded Rectangle
      Top             =   5130
      Visible         =   0   'False
      Width           =   3435
   End
End
Attribute VB_Name = "frmFechamento_Encerrantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String
Dim strNome_label As String
Dim intcontbomba As Integer
Dim intcontBomba_bico As Integer
Dim rstBombas As New ADODB.Recordset
Dim rstBomba_bico As New ADODB.Recordset
Public strCasas_Decimais As String
Public strNumero_PDV As String
Public strOperador As String
'Conexões-----------------------------------------------------
Dim CNconexao As New DLLConexao_Sistema.conexao
Dim CNconexao_local_pdv As New DLLConexao_Sistema.conexao
'-------------------------------------------------------------
Option Explicit

Private Sub cmdOk_Click()

    Dim intContador_bicos As Integer
    Dim rstEncerante As New ADODB.Recordset
    Dim lngID_Encerrante As Long
    Dim IDBomba_bico As Long
    
    'Verificando se existe encerrantes não preenchidos
    Do While txtBico.Count > intContador_bicos
       intContador_bicos = intContador_bicos + 1
       If Me.txtBico.Item(intContador_bicos).Visible = True Then
          If Me.txtBico.Item(intContador_bicos).Text = "" Or txtBico.Item(intContador_bicos).Text = "0" Then
             MsgBox "Bico com valor não informado!Verifique.", vbCritical, "Only Tech"
             Me.txtBico.Item(intContador_bicos).SetFocus
             GoTo Fim_operacao
          End If
       End If
    Loop
    
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       'Abrindo uma conexão nova com o Retaguarda
       If CNconexao.CNconexao <> "" Then
          Set CNconexao = Nothing
          CNconexao.Banco = "BDRetaguarda"
          CNconexao.Abrir_conexao "Otica"
       Else
          CNconexao.Banco = "BDRetaguarda"
          CNconexao.Abrir_conexao "Otica"
       End If
    End If
    
'''    'Abrindo uma conexão nova com o Retaguarda
'''    If CNconexao_local_pdv.CNconexao <> "" Then
'''       Set CNconexao_local_pdv = Nothing
'''       CNconexao_local_pdv.Banco = "BDPDV"
'''       CNconexao_local_pdv.Abrir_conexao "PDV"
'''    Else
'''      CNconexao_local_pdv.Banco = "BDPDV"
'''      CNconexao_local_pdv.Abrir_conexao "PDV"
'''    End If
'''
    strSql = Empty
    strSql = "INSERT INTO TBENCERRANTE (FKCodigo_TBPdv,FKCodigo_TBOperadores_ecf,DFData_TBEncerrante,DFHora_TBEncerrante,DFAbertura_fechamento_TBEncerrante) " & _
             "VALUES( " & frmTela_Venda.txtNumero_check_out.Text & ", " & frmTela_Venda.strCodigo_Operador & ", '" & Format(frmTela_Venda.dtpData_operacao, "YYYYMMDD") & "', '" & Format(Now, "HH:MM:SS") & "','1')"
             
    On Error GoTo Erro
    
    'Gravando a TBEncerrantes
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       'Abrindo a Transação
       CNconexao.CNconexao.BeginTrans
       'Gravando a TBEncerrantes
       CNconexao.CNconexao.Execute strSql
    End If
    
    'Comitando a transação
    CNconexao.CNconexao.CommitTrans
    
    'Pegando o ID do encerrante
    strSql = Empty
    strSql = "SELECT max(PKId_TBEncerrante)ID_Encerrante FROM TBENCERRANTE"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstEncerante, "Otica", Me
    
    lngID_Encerrante = rstEncerante!ID_Encerrante
    Set rstEncerante = Nothing
    
''    'Abrindo a Transação
''    'CNconexao_local_pdv.CNconexao.BeginTrans
''    'Gravando a TBEncerrantes
''    'CNconexao_local_pdv.CNconexao.Execute strSql
''
    'Gravando a TBEncerrantes Bico
    On Error GoTo Erro_Encerrantes_Bico
    
    'Abrindo +1 Transação
    CNconexao.CNconexao.BeginTrans
    
    intContador_bicos = 0
    
    Do While txtBico.Count > intContador_bicos
       intContador_bicos = intContador_bicos + 1
       If Me.txtBico.Item(intContador_bicos).Visible = True Then
          'Localizando bomba deste bico
          IDBomba_bico = Funcoes_Gerais.Localiza_ID("PKId_TBBomba_bico", "IXCodigo_TBBomba_bico", Trim(Replace(lblBico.Item(intContador_bicos).Caption, "BICO", "")), "TBBOMBA_BICO", "Otica", Me)
          
          strSql = Empty
          strSql = "INSERT INTO TBENCERRANTE_BOMBA (FKId_TBBomba_bico,FKId_TBEncerrante,DFEncerrante_TBEncerrante_Bomba) " & _
                   "VALUES( " & IDBomba_bico & ", " & lngID_Encerrante & ", " & Funcoes_Gerais.Grava_Moeda(Me.txtBico.Item(intContador_bicos).Text) & ")"
        
          If frmTela_Venda.booIntegracao_Retaguarda = True Then
             'Gravando a TBEncerrantes
             CNconexao.CNconexao.Execute strSql
          End If
          
       End If
    Loop
    
    frmTela_Venda.booConsulta = False
    
    'Comitando a transação
    CNconexao.CNconexao.CommitTrans
    
    'Gravando a TBEncerrantes Bico
    On Error GoTo Erro_bomba_bico
    
    'Abrindo +1 Transação
    CNconexao.CNconexao.BeginTrans
    
    intContador_bicos = 0
    
    Do While txtBico.Count > intContador_bicos
       intContador_bicos = intContador_bicos + 1
       If Me.txtBico.Item(intContador_bicos).Visible = True Then
          'Gravando o ultimo encerrante na tabela de bicos
          strSql = Empty
          strSql = "UPDATE TBBOMBA_BICO " & _
                   "SET DFUltimo_encerrante_TBBomba_bico = " & Funcoes_Gerais.Grava_Moeda(Me.txtBico.Item(intContador_bicos).Text) & " WHERE IXCodigo_TBBomba_bico = " & Trim(Replace(lblBico.Item(intContador_bicos).Caption, "BICO", "")) & ""
        
          If frmTela_Venda.booIntegracao_Retaguarda = True Then
             'Gravando a TBEncerrantes
             CNconexao.CNconexao.Execute strSql
          End If
       End If
    Loop
    
    'Comitando a transação
    CNconexao.CNconexao.CommitTrans
    
    Set CNconexao = Nothing
''  Set CNconexao_local_pdv = Nothing

    MsgBox "Fechamento de encerrantes efetuado com sucesso!", vbInformation, "Only Tech"
    
    Unload Me
Fim_operacao:
    Exit Sub
    
Erro:
    'Roolback na transação
    CNconexao.CNconexao.RollbackTrans
    Set CNconexao = Nothing
    
    MsgBox "Ocorreu um erro: " & Err.Number & "-" & Err.Description, vbCritical, "Only Tech"
    
Erro_Encerrantes_Bico:

    'Roolback na transação
    CNconexao.CNconexao.RollbackTrans
    
    'Comitando a transação
    CNconexao.CNconexao.BeginTrans
    
    strSql = Empty
    strSql = "DELETE FROM TBENCERRANTE WHERE PKId_TBEncerrante = " & lngID_Encerrante & ""
    CNconexao.CNconexao.Execute strSql
    
    'Comitando a transação
    CNconexao.CNconexao.CommitTrans
    
    Set CNconexao = Nothing
    
    MsgBox "Ocorreu um erro: " & Err.Number & "-" & Err.Description, vbCritical, "Only Tech"
    
Erro_bomba_bico:

    'Roolback na transação
    CNconexao.CNconexao.RollbackTrans
    
    'Comitando a transação
    CNconexao.CNconexao.BeginTrans
    
    strSql = Empty
    strSql = "DELETE FROM TBENCERRANTE_BOMBA WHERE FKId_TBBomba_bico = " & lngID_Encerrante & ""
    CNconexao.CNconexao.Execute strSql
    
    'Comitando a transação
    CNconexao.CNconexao.CommitTrans
    
    Set CNconexao = Nothing
    
    MsgBox "Ocorreu um erro: " & Err.Number & "-" & Err.Description, vbCritical, "Only Tech"
    
     
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Call Monta_Tela
End Sub

Private Sub txtBico_GotFocus(Index As Integer)
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub
Private Function Monta_Tela()
    
    'Operador
    lblOperador.Caption = "Operador: " & frmTela_Venda.strOperador
    Me.lblPDV.Caption = "N° PDV: " & frmTela_Venda.txtNumero_check_out
    
    Dim intBombas As Integer
    
    'Buscando as bombas
    strSql = Empty
    strSql = "SELECT * FROM TBBomba"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstBombas, "Otica", Me
    
    If rstBombas.BOF = True And rstBombas.EOF = True Then
       MsgBox "Bombas não cadastradas!Verifique.", vbCritical, "Only Tech"
       Set rstBombas = Nothing
       End
    End If
    
    rstBombas.MoveFirst
    intBombas = rstBombas.RecordCount
    
    Do While rstBombas.EOF = False

       intcontbomba = intcontbomba + 1
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       'BOMBAS
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       lblBomba.Item(intcontbomba).Caption = "BOMBA " & rstBombas!PKId_TBBomba & " - " & rstBombas!DFDescricao_TBBomba
       lblBomba.Item(intcontbomba).Visible = True
       Me.Shape_bomba.Item(intcontbomba).Visible = True
       Me.Shape_bomba2.Item(intcontbomba).Visible = True
       
       'Buscando os bicos
       strSql = Empty
       strSql = "SELECT * FROM TBBomba_bico WHERE FKId_TBBomba = " & rstBombas!PKId_TBBomba & ""
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstBomba_bico, "Otica", Me
       
       rstBomba_bico.MoveFirst
       Dim intBicos_bomba As Integer
       
       intBicos_bomba = rstBomba_bico.RecordCount
       
       Do While rstBomba_bico.EOF = False
          intcontBomba_bico = intcontBomba_bico + 1
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          'BOMBAS
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          lblBico.Item(intcontBomba_bico).Caption = "BICO " & rstBomba_bico!IXCodigo_TBBomba_bico
          lblBico.Item(intcontBomba_bico).Visible = True
          Me.txtBico.Item(intcontBomba_bico).Visible = True
          Me.Shape_bico.Item(intcontBomba_bico).Visible = True
          Me.txtBico.Item(intcontBomba_bico).Text = 0
          If frmTela_Venda.strCasas_Decimais = 2 Then
               txtBico.Item(intcontBomba_bico).Text = Format(txtBico.Item(intcontBomba_bico), "#,###0.00")
          End If
          If frmTela_Venda.strCasas_Decimais = 3 Then
               txtBico.Item(intcontBomba_bico).Text = Format(txtBico.Item(intcontBomba_bico), "#,###0.000")
          End If
          rstBomba_bico.MoveNext
       Loop
       
       'Limpando os componentes de tela em excesso
       If intBicos_bomba < 4 Then
       
          Dim intbicos_sobra As Integer
          
          intbicos_sobra = 4 - intBicos_bomba
          
          Do While intbicos_sobra > 0
             intcontBomba_bico = intcontBomba_bico + 1
             lblBico.Item(intcontBomba_bico).Visible = False
             intbicos_sobra = intbicos_sobra - 1
          Loop
          
       End If
       
       Set rstBomba_bico = Nothing
       rstBombas.MoveNext
    Loop
    
    'Limpando os componentes de tela em excesso
    If intBombas < 4 Then
    
       Dim intbombas_sobra As Integer
        
       intbombas_sobra = 6 - intcontbomba
         
       Do While intbombas_sobra > 0
          intcontbomba = intcontbomba + 1
          lblBomba.Item(intcontbomba).Visible = False
          intbombas_sobra = intbombas_sobra - 1
       Loop
          
    End If
       
    Set rstBombas = Nothing
    Set rstBomba_bico = Nothing

End Function

Private Sub txtBico_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 44 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtBico_LostFocus(Index As Integer)
     If txtBico.Item(Index).Text <> "" Then
        If txtBico.Item(Index).Text = "," Then
           txtBico.Item(Index).Text = 0
        End If
        If txtBico.Item(Index).Text = 0 Then
           txtBico.Item(Index).Text = Empty
           txtBico.Item(Index).SetFocus
        Else
            If frmTela_Venda.strCasas_Decimais = 2 Then
               txtBico.Item(Index).Text = Format(txtBico.Item(Index), "#,###0.00")
            End If
            If frmTela_Venda.strCasas_Decimais = 3 Then
               txtBico.Item(Index).Text = Format(txtBico.Item(Index), "#,###0.000")
            End If
        End If
     End If
End Sub
