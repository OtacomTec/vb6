VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VB6 Sample Code - Printing and Magnetic Encoding"
   ClientHeight    =   6120
   ClientLeft      =   9825
   ClientTop       =   5205
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   8865
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   7800
      TabIndex        =   12
      Top             =   5520
      Width           =   900
   End
   Begin VB.CommandButton btnSubmit 
      Caption         =   "&Submit"
      Height          =   495
      Left            =   6720
      TabIndex        =   11
      Top             =   5520
      Width           =   900
   End
   Begin VB.Frame frameVersions 
      Caption         =   "DLL Versions:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8675
      Begin VB.Label lblVersions 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   400
         Width           =   825
      End
   End
   Begin VB.Frame frameMag 
      Caption         =   "Magnetic Encoder:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   8675
      Begin VB.CheckBox cbMag 
         Caption         =   "Check"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   450
         Width           =   1095
      End
   End
   Begin VB.Frame framePrint 
      Caption         =   "Printer Selections:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   8675
      Begin VB.CheckBox cbBack 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox cbFront 
         Caption         =   "Front"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   450
         Width           =   1095
      End
   End
   Begin VB.Frame frameStatus 
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   8675
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   405
         Width           =   60
      End
   End
   Begin VB.Frame framePrn 
      Caption         =   "Printers:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8675
      Begin VB.ComboBox cboPrn 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'* CONFIDENTIAL AND PROPRIETARY
'*
'* The source code and other information contained herein is the confidential and the exclusive property of
'* ZIH Corp. and is subject to the terms and conditions in your end user license agreement.
'* This source code, and any other information contained herein, shall not be copied, reproduced, published,
'* displayed or distributed, in whole or in part, in any medium, by any means, for any purpose except as
'* expressly permitted under such license agreement.
'*
'* Copyright ZIH Corp. 2010
'*
'* ALL RIGHTS RESERVED
'***********************************************
'File: frmMain.frm
'Description: Form allowing users to select printing options and print.
'$Revision: 1 $
'$Date: 2010/12/13 $
'*******************************************************************************/

Option Explicit

' Local Variables -------------------------------------------------------------------------------------------

Private graphicSDKVersion   As String
Private prnSDKVersion       As String

' Load Form Event -------------------------------------------------------------------------------------------

Private Sub Form_Load()
    
    cboPrnInit
    FormConfig
   
End Sub

' Form Configuration ----------------------------------------------------------------------------------------

Private Sub FormConfig()

    Dim msg As String
    
    On Error GoTo FormConfig_Error
    
    msg = ""
    
    ' Gets graphics dll version
    '     and if present enables the print frame
    
    GetGraphicsDllVersion graphicSDKVersion
    
    If graphicSDKVersion <> "" Then
        Me.framePrint.Enabled = True
        msg = "Graphics: " & graphicSDKVersion & "; "
    End If
        
    ' Gets printer dll version
    '     and if present enables the magnetic encoding frame
    
    GetPrinterDllVersion prnSDKVersion

    If prnSDKVersion <> "" Then
        Me.frameMag.Enabled = True
        msg = msg & "Printer: " & prnSDKVersion & "; "
    End If
    
    ' Displays dll versions
    
    Me.lblVersions = msg
    
FormConfig_Exit:
    On Error GoTo 0
    Exit Sub
    
FormConfig_Error:
    MsgBox "Error in FormConfig: " & Err.Description
    GoTo FormConfig_Exit
End Sub

' Initializes Printer combo box -----------------------------------------------------------------------------

Private Sub cboPrnInit()

    On Error GoTo cboPrnInit_Error
    
    Dim p As Printer
    
    For Each p In Printers
        Me.cboPrn.AddItem p.DeviceName
    Next
    
cboPrnInit_Exit:
    On Error GoTo 0
    Exit Sub
    
cboPrnInit_Error:
    MsgBox "Error in cboPrnInit: " & Err.Description
    GoTo cboPrnInit_Exit
End Sub

' Button : Exit Application --------------------------------------------------------------------------------

Private Sub btnExit_Click()

    Unload Me

End Sub

' Button : Runs selected operations -------------------------------------------------------------------------

Private Sub btnSubmit_Click()

    Dim msg         As String
    
    On Error GoTo btnSubmit_Click_Error
    
    ' Verifies that a printer has been selected
    
    If Me.cboPrn.text = "" Then
        msg = "Error: A printer has not been selected"
        GoTo btnSubmit_Click_Exit
    End If
    
    ' Verifies that at least one selection has been made
    
    If Me.cbBack.Value = 0 And Me.cbFront.Value = 0 And Me.cbMag.Value = 0 Then
        msg = "Error: No selections have been made"
        GoTo btnSubmit_Click_Exit
    End If
    
    ' Magnetic Encoding
        
    Dim eject       As Boolean

    If Me.cbMag.Value <> 0 Then
    
        ' ejects the card after magnetic encoding if neither front or back printing is selected
        
        eject = IIf(Me.cbBack.Value <> 0 Or Me.cbFront.Value <> 0, False, True)
        
        ' Encodes and verifies all 3 tracks
        
        ' Note that we can only encode 6 characters or less to track 2
        ' when using printer firmware version lower than 2.00.03
        
        MagCode Me.cboPrn.text, "ABCDEFGH", "123456", "87654321", eject, msg
        If msg <> "" Then
            GoTo btnSubmit_Click_Exit
        End If
        
    End If
    
    ' Path and filename of the image to printed
    
    Dim filename As String
    filename = App.Path & "\" & "Zebra.bmp"
    
    If Me.cbFront.Value <> 0 And Me.cbBack.Value = 0 Then
        PrintFrontSideOnly Me.cboPrn.text, "Front Side Text", filename, msg
        If msg = "" Then Me.lblStatus.Caption = "No Error : Printing Front Side Only"
    
    ElseIf Me.cbFront.Value <> 0 And Me.cbBack.Value <> 0 Then
        PrintBothSides Me.cboPrn.text, "Front Side Text", filename, "Back Side Text", msg
        If msg = "" Then Me.lblStatus.Caption = "No Error : Printing Both Sides"
    
    End If
        
btnSubmit_Click_Exit:
    If msg <> "" Then
        Me.lblStatus = msg
    Else
        Me.lblStatus = "No Errors"
    End If
    
    On Error GoTo 0
    Exit Sub
    
btnSubmit_Click_Error:
    MsgBox "Error in btnSubmit_Click: " & Err.Description
    GoTo btnSubmit_Click_Exit
End Sub

' Combo Boxes -----------------------------------------------------------------------------------------------

Private Sub cbBack_Click()
    If Me.cbBack.Value <> 0 Then Me.cbFront.Value = 1
End Sub

Private Sub cbFront_Click()
    If Me.cbFront.Value = 0 Then Me.cbBack.Value = 0
End Sub

