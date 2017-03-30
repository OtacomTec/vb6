VERSION 5.00
Object = "*\AVertMenu.vbp"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Vertical Menu Demonstration - View KnowledgeBase Articles"
   ClientHeight    =   6615
   ClientLeft      =   1800
   ClientTop       =   2595
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   705
   Begin VertMenu.VerticalMenu VerticalMenu1 
      Height          =   6225
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   10980
      MenusMax        =   4
      MenuCaption1    =   "Fixes"
      MenuItemsMax1   =   7
      MenuItemIcon11  =   "Test.frx":0000
      MenuItemCaption11=   "Q100190"
      MenuItemIcon12  =   "Test.frx":031A
      MenuItemCaption12=   "Q100367"
      MenuItemIcon13  =   "Test.frx":0634
      MenuItemCaption13=   "Q101257"
      MenuItemIcon14  =   "Test.frx":094E
      MenuItemCaption14=   "Q113281"
      MenuItemIcon15  =   "Test.frx":0C68
      MenuItemCaption15=   "Q74517"
      MenuItemIcon16  =   "Test.frx":0F82
      MenuItemCaption16=   "Q84475"
      MenuItemIcon17  =   "Test.frx":129C
      MenuItemCaption17=   "Q93436"
      MenuCaption2    =   "Problems"
      MenuItemsMax2   =   5
      MenuItemIcon21  =   "Test.frx":15B6
      MenuItemCaption21=   "Q79094"
      MenuItemIcon22  =   "Test.frx":18D0
      MenuItemCaption22=   "Q79599"
      MenuItemIcon23  =   "Test.frx":1BEA
      MenuItemCaption23=   "Q82157"
      MenuItemIcon24  =   "Test.frx":1F04
      MenuItemCaption24=   "Q80645"
      MenuItemIcon25  =   "Test.frx":221E
      MenuItemCaption25=   "Q84483"
      MenuCaption3    =   "Bugs"
      MenuItemsMax3   =   6
      MenuItemIcon31  =   "Test.frx":2538
      MenuItemCaption31=   "Q100193"
      MenuItemIcon32  =   "Test.frx":2852
      MenuItemCaption32=   "Q115779"
      MenuItemIcon33  =   "Test.frx":2B6C
      MenuItemCaption33=   "Q139567"
      MenuItemIcon34  =   "Test.frx":2E86
      MenuItemCaption34=   "Q145618"
      MenuItemIcon35  =   "Test.frx":31A0
      MenuItemCaption35=   "Q76520"
      MenuItemIcon36  =   "Test.frx":34BA
      MenuItemCaption36=   "Q77393"
      MenuCaption4    =   "Programs"
      MenuItemsMax4   =   4
      MenuItemIcon41  =   "Test.frx":37D4
      MenuItemCaption41=   "Calculator"
      MenuItemIcon42  =   "Test.frx":3AEE
      MenuItemCaption42=   "Clock"
      MenuItemIcon43  =   "Test.frx":3E08
      MenuItemCaption43=   "Notepad"
      MenuItemIcon44  =   "Test.frx":4122
      MenuItemCaption44=   "Paint"
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6225
      Left            =   1755
      TabIndex        =   1
      Top             =   180
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   10980
      _Version        =   327680
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"Test.frx":443C
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub VerticalMenu1_MenuItemClick(MenuNumber As Long, MenuItem As Long)
    Dim Path As String
    
    On Error Resume Next
    Path = App.Path & "\"
    Select Case MenuNumber
        Case 1
            Select Case MenuItem
                Case 1
                    RichTextBox1.filename = Path & "Q100190.rtf"
                Case 2
                    RichTextBox1.filename = Path & "Q100367.rtf"
                Case 3
                    RichTextBox1.filename = Path & "Q101257.rtf"
                Case 4
                    RichTextBox1.filename = Path & "Q113281.rtf"
                Case 5
                    RichTextBox1.filename = Path & "Q74517.rtf"
                Case 6
                    RichTextBox1.filename = Path & "Q84475.rtf"
                Case 7
                    RichTextBox1.filename = Path & "Q93436.rtf"
            End Select
        Case 2
            Select Case MenuItem
                Case 1
                    RichTextBox1.filename = Path & "Q79094.rtf"
                Case 2
                    RichTextBox1.filename = Path & "Q79599.rtf"
                Case 3
                    RichTextBox1.filename = Path & "Q82157.rtf"
                Case 4
                    RichTextBox1.filename = Path & "Q80645.rtf"
                Case 5
                    RichTextBox1.filename = Path & "Q84483.rtf"
            End Select
        Case 3
            Select Case MenuItem
                Case 1
                    RichTextBox1.filename = Path & "Q100193.rtf"
                Case 2
                    RichTextBox1.filename = Path & "Q115779.rtf"
                Case 3
                    RichTextBox1.filename = Path & "Q139567.rtf"
                Case 4
                    RichTextBox1.filename = Path & "Q145618.rtf"
                Case 5
                    RichTextBox1.filename = Path & "Q76520.rtf"
                Case 6
                    RichTextBox1.filename = Path & "Q77393.rtf"
            End Select
        Case 4
            Select Case MenuItem
                Case 1
                    Shell "calc.exe", vbNormalFocus
                Case 2
                    Shell "clock.exe", vbNormalFocus
                Case 3
                    Shell "notepad.exe", vbNormalFocus
                Case 4
                    Shell "mspaint.exe", vbNormalFocus
            End Select
    End Select
End Sub
