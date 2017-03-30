VERSION 5.00
Begin VB.Form formConvertToPDF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text-PDF v1.0"
   ClientHeight    =   6210
   ClientLeft      =   2625
   ClientTop       =   1395
   ClientWidth     =   6630
   Icon            =   "formPDF.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6630
   Begin VB.ComboBox cmbPageSize 
      Height          =   315
      ItemData        =   "formPDF.frx":08CA
      Left            =   3540
      List            =   "formPDF.frx":08D7
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   5040
      Width           =   1065
   End
   Begin VB.ComboBox cmbFontSize 
      Height          =   315
      ItemData        =   "formPDF.frx":08FE
      Left            =   2160
      List            =   "formPDF.frx":0917
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   5040
      Width           =   690
   End
   Begin VB.ComboBox cmbRotation 
      Height          =   315
      ItemData        =   "formPDF.frx":0944
      Left            =   2835
      List            =   "formPDF.frx":0954
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   5040
      Width           =   690
   End
   Begin VB.ComboBox cmbFont 
      Height          =   315
      ItemData        =   "formPDF.frx":096D
      Left            =   120
      List            =   "formPDF.frx":097A
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Frame frmSep 
      Height          =   120
      Left            =   -60
      TabIndex        =   23
      Top             =   5385
      Width           =   6765
   End
   Begin VB.CommandButton btnClose 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   420
      Left            =   5190
      TabIndex        =   21
      Top             =   5580
      Width           =   1350
   End
   Begin VB.Frame frmTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   -45
      TabIndex        =   20
      Top             =   -15
      Width           =   6720
      Begin VB.Label lblCaption 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Converter Texto  para PDF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   210
         TabIndex        =   22
         Top             =   225
         Width           =   5025
      End
      Begin VB.Image imgIcon 
         Height          =   630
         Left            =   5715
         Picture         =   "formPDF.frx":099B
         Stretch         =   -1  'True
         Top             =   135
         Width           =   570
      End
   End
   Begin VB.TextBox txtTitle 
      Height          =   360
      Left            =   1470
      TabIndex        =   9
      Top             =   3045
      Width           =   5100
   End
   Begin VB.CommandButton btnConvert 
      Caption         =   "&Gerar PDF"
      Default         =   -1  'True
      Height          =   420
      Left            =   3750
      TabIndex        =   19
      Top             =   5580
      Width           =   1350
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Procurar."
      Height          =   330
      Left            =   4980
      TabIndex        =   15
      Top             =   5010
      Width           =   1560
   End
   Begin VB.TextBox txtOutputFile 
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   4650
      Width           =   6420
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "&Procurar."
      Height          =   330
      Left            =   4980
      TabIndex        =   12
      Top             =   4140
      Width           =   1560
   End
   Begin VB.TextBox txtFilename 
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   3780
      Width           =   6420
   End
   Begin VB.TextBox txtSubject 
      Height          =   360
      Left            =   1470
      TabIndex        =   7
      Top             =   2565
      Width           =   5100
   End
   Begin VB.TextBox txtKeywords 
      Height          =   360
      Left            =   1470
      TabIndex        =   5
      Top             =   2100
      Width           =   5100
   End
   Begin VB.TextBox txtCreator 
      Height          =   360
      Left            =   1470
      TabIndex        =   3
      Top             =   1635
      Width           =   5100
   End
   Begin VB.TextBox txtAuthor 
      Height          =   360
      Left            =   1470
      TabIndex        =   1
      Top             =   1185
      Width           =   5100
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Titulo:"
      Height          =   255
      Left            =   150
      TabIndex        =   8
      Top             =   3090
      Width           =   1230
   End
   Begin VB.Label lblOutputFile 
      Caption         =   "Arq. SaÌda :"
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   4425
      Width           =   2895
   End
   Begin VB.Label lblFilename 
      Caption         =   "Arq. Entrada:"
      Height          =   240
      Left            =   135
      TabIndex        =   10
      Top             =   3525
      Width           =   2895
   End
   Begin VB.Label lblSubject 
      Alignment       =   1  'Right Justify
      Caption         =   "Assunto:"
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   2610
      Width           =   1230
   End
   Begin VB.Label lblKeyword 
      Alignment       =   1  'Right Justify
      Caption         =   "Palavras Chave:"
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   2145
      Width           =   1230
   End
   Begin VB.Label lblCreator 
      Alignment       =   1  'Right Justify
      Caption         =   "Autor:"
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   1680
      Width           =   1230
   End
   Begin VB.Label lblAuthor 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome"
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   1230
      Width           =   1230
   End
End
Attribute VB_Name = "formConvertToPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
    
Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Private Const OFN_EXPLORER = &H80000 ' new look commdlg
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0


Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Dim Position As Long
Dim pageNo As Long
Dim lineNo As Long
Dim pageHeight As Long
Dim pageWidth As Long
Dim location(1 To 5000) As Long
Dim pageObj(1 To 5000) As Long
Dim lines As Long
Dim obj As Long
Dim Tpages As Long
Dim encoding As Long
Dim resources As Long
Dim pages As Variant
Dim author As String
Dim creator As String
Dim keywords As String
Dim subject As String
Dim Title As String
Dim BaseFont As String
Dim pointSize As Currency
Dim vertSpace As Currency
Dim rotate As Integer
Dim info As Long
Dim root As Long
Dim npagex As Double
Dim npagey As Long
Dim filetxt As String
Dim filepdf As String
Dim linelen As Long
Dim cache As String
Dim cmdline As String

Const AppName = "Text-PDF v1.0"

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
  txtCreator.Text = AppName
  cmbFont.ListIndex = 1
  cmbFontSize.ListIndex = 1
  cmbRotation.ListIndex = 0
  cmbPageSize.ListIndex = 0

  cmdline = LCase(Command)
  If cmdline Like """*""" Then
    cmdline = Mid(cmdline, 2, Len(cmdline) - 2)
  End If
  
  If FileExists(cmdline) Then
    txtFilename.Text = cmdline
    txtOutputFile.Text = Left(cmdline, Len(cmdline) - 4) & ".pdf"
    btnConvert_Click
  End If
End Sub

Private Sub lblCredits_Click(Index As Integer)
  On Error Resume Next
  ShellExecute 0, vbNullString, lblCredits(Index).ToolTipText, vbNullString, vbNullString, 1
End Sub

Private Sub txtAuthor_GotFocus()
  txtAuthor.SelStart = 0
  txtAuthor.SelLength = Len(txtAuthor.Text)
End Sub

Private Sub txtCreator_GotFocus()
  txtCreator.SelStart = 0
  txtCreator.SelLength = Len(txtCreator.Text)
End Sub

Private Sub txtSubject_GotFocus()
  txtSubject.SelStart = 0
  txtSubject.SelLength = Len(txtSubject.Text)
End Sub

Private Sub txtTitle_GotFocus()
  txtTitle.SelStart = 0
  txtTitle.SelLength = Len(txtTitle.Text)
End Sub

Private Sub txtKeywords_GotFocus()
  txtKeywords.SelStart = 0
  txtKeywords.SelLength = Len(txtKeywords.Text)
End Sub

Private Sub txtFilename_GotFocus()
  txtFilename.SelStart = 0
  txtFilename.SelLength = Len(txtFilename.Text)
End Sub

Private Sub txtOutputFile_GotFocus()
  txtOutputFile.SelStart = 0
  txtOutputFile.SelLength = Len(txtOutputFile.Text)
End Sub

Private Sub btnClose_Click()
  Unload Me
End Sub

Private Sub btnOpen_Click()
  Dim filename As String
  On Local Error Resume Next
  filename = OpenDialog(Me, "Text files (*.txt)|*.txt|All files (*.*)|*.*", _
                   "Select a text file", "")
  If Len(filename) Then
    txtFilename.Text = filename
    filename = txtFilename.Text
    txtOutputFile.Text = Left(filename, Len(filename) - 3) & "pdf"
  End If
End Sub

Private Sub btnSave_Click()
  Dim filename As String
  On Local Error Resume Next
  filename = SaveDialog(Me, "Portable Document Format files (*.pdf)|*.pdf", _
                        "Save PDF As", "", "")
  If Len(filename) Then
    txtOutputFile.Text = filename
  End If
End Sub

Private Sub btnSource_Click()
  On Local Error Resume Next
End Sub

Private Sub btnConvert_Click()
  If txtFilename.Text <> "" And txtOutputFile.Text <> "" Then
    ConvertToPDF txtFilename.Text, txtOutputFile.Text, _
                 txtAuthor.Text, txtCreator.Text, txtKeywords.Text, _
                 txtSubject.Text, txtTitle.Text, _
                 cmbFont.Text, Val(cmbFontSize.Text), Val(cmbRotation.Text), _
                 Val(cmbPageSize.Text), Val(Right(cmbPageSize.Text, 3))
    If FileExists(cmdline) Then
      Unload Me
    ElseIf MsgBox("PDF foi criado." & vbCr & vbCr & "Deseja abrir o arquivo PDF gerado ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
      ShellExecute 0, vbNullString, txtOutputFile.Text, vbNullString, vbNullString, 1
    End If
  Else
    MsgBox "Informe o nome do arquivo."
  End If
End Sub

Public Sub ConvertToPDF(filename As String, outputfile As String, _
                        Optional TextAuthor As String, Optional TextCreator As String, Optional TextKeywords As String, _
                        Optional TextSubject As String, Optional TextTitle As String, _
                        Optional FontName As String = "Courier", Optional FontSize As Integer = 10, Optional Rotation As Integer, _
                        Optional pwidth As Single = 8.5, Optional pheight As Single = 11)
  On Error GoTo er
  If Not FileExists(filename) Then
    MsgBox "Arquivo '" & filename & "' n„o existe."
    Exit Sub
  ElseIf FileExists(outputfile) Then
    Kill outputfile
  End If

  initialize FontName, FontSize, Rotation, pwidth, pheight
  
  author = TextAuthor
  creator = TextCreator
  keywords = TextKeywords
  subject = TextSubject
  Title = TextTitle
  filetxt = filename
  filepdf = outputfile
  
  Call WriteStart
  Call WriteHead
  Call WritePages
  Call endpdf
  Exit Sub
er:
  MsgBox Err.Description
End Sub

Private Sub initialize(FontName As String, FontSize As Integer, Rotation As Integer, pwidth As Single, pheight As Single)
  pageHeight = 72 * pheight
  pageWidth = 72 * pwidth

  BaseFont = FontName ' Courier, Times-Roman, Arial
  pointSize = FontSize ' Font Size; n„o altere
  vertSpace = FontSize * 1.2 ' Vertical spacing
  rotate = Rotation ' degrees to rotate; try setting 90,180,etc
  lines = (pageHeight - 72) / vertSpace ' no of lines on one page
  
  Select Case LCase(FontName)
   Case "courier": linelen = 1.5 * pageWidth / pointSize
   Case "arial": linelen = 2 * pageWidth / pointSize
  'Case "Times-Roman": linelen = 2.2 * pageWidth / pointSize
   Case Else: linelen = 2.2 * pageWidth / pointSize
  End Select

  obj = 0
  npagex = pageWidth / 2
  npagey = 25
  pageNo = 0
  Position = 0
  cache = ""
End Sub

Private Sub writepdf(stre As String, Optional flush As Boolean)
  On Local Error Resume Next
  Position = Position + Len(stre)
  cache = cache & stre & vbCr
  If Len(cache) > 32000 Or flush Then
    Open filepdf For Append As #1
    Print #1, cache;
    Close #1
    cache = ""
  End If
End Sub
  
Private Sub WriteStart()
  writepdf ("%PDF-1.2")
  writepdf ("%‚„œ”")
End Sub

Private Sub WriteHead()
  Dim CreationDate As String
  On Error GoTo er
    CreationDate = "D:" & Format(Now, "YYYYMMDDHHNNSS")
    obj = obj + 1
    location(obj) = Position
    info = obj
    
    writepdf (obj & " 0 obj")
    writepdf ("<<")
    writepdf ("/Author (" & author & ")")
    writepdf ("/CreationDate (" & CreationDate & ")")
    writepdf ("/Creator (" & creator & ")")
    writepdf ("/Producer (" & AppName & ")")
    writepdf ("/Title (" & Title & ")")
    writepdf ("/Subject (" & subject & ")")
    writepdf ("/Keywords (" & keywords & ")")
    writepdf (">>")
    writepdf ("endobj")
    
    obj = obj + 1
    root = obj
    obj = obj + 1
    Tpages = obj
    encoding = obj + 2
    resources = obj + 3
    
    obj = obj + 1
    location(obj) = Position
    writepdf (obj & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Font")
    writepdf ("/Subtype /Type1")
    writepdf ("/Name /F1")
    writepdf ("/Encoding " & encoding & " 0 R")
    writepdf ("/BaseFont /" & BaseFont)
    writepdf (">>")
    writepdf ("endobj")
    
    obj = obj + 1
    location(obj) = Position
    writepdf (obj & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Encoding")
    writepdf ("/BaseEncoding /WinAnsiEncoding")
    writepdf (">>")
    writepdf ("endobj")
    
    obj = obj + 1
    location(obj) = Position
    writepdf (obj & " 0 obj")
    writepdf ("<<")
    writepdf ("  /Font << /F1 " & obj - 2 & " 0 R >>")
    writepdf ("  /ProcSet [ /PDF /Text ]")
    writepdf (">>")
    writepdf ("endobj")
  Exit Sub
er:
  MsgBox Err.Description
End Sub
  
Private Sub WritePages()
  Dim i As Integer
  Dim line As String, tmpline As String, beginstream As String
  On Error GoTo er
    Open filetxt For Input As #2
      beginstream = StartPage
      lineNo = -1
      Do Until EOF(2)
        Line Input #2, line
        lineNo = lineNo + 1
        
        'quebra de p·gina
        If lineNo >= lines Or InStr(line, Chr(12)) > 0 Then
          writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
          writepdf ("(" & pageNo & ") Tj")
          writepdf ("/F1 " & pointSize & " Tf")
          endpage (beginstream)
          beginstream = StartPage
        End If
        
        line = ReplaceText(ReplaceText(line, "(", "\("), ")", "\)")
        line = Trim(line)
        
        If Len(line) > linelen Then
          
          'quebra de linha
          Do While Len(line) > linelen
            tmpline = Left(line, linelen)
            For i = Len(tmpline) To Len(tmpline) \ 2 Step -1
              If InStr("*&^%$#,. ;<=>[])}!""", Mid(tmpline, i, 1)) Then
                tmpline = Left(tmpline, i)
                Exit For
              End If
            Next
            
            line = Mid$(line, Len(tmpline) + 1)
            writepdf ("T* (" & tmpline & vbCrLf & ") Tj")
            lineNo = lineNo + 1
            
            'quebra de p·gina
            If lineNo >= lines Or InStr(line, Chr(12)) > 0 Then
              writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
              writepdf ("(" & pageNo & ") Tj")
              writepdf ("/F1 " & pointSize & " Tf")
              endpage (beginstream)
              beginstream = StartPage
            End If
          Loop
          
          lineNo = lineNo + 1
          writepdf ("T* (" & line & vbCrLf & ") Tj")
        
        Else
          
          writepdf ("T* (" & line & vbCrLf & ") Tj")
        
        End If
      Loop
    Close #2
    writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
    writepdf ("(" & pageNo & ") Tj")
    writepdf ("/F1 " & pointSize & " Tf")
    endpage (beginstream)
  Exit Sub
er:
  MsgBox Err.Description
  Close
End Sub

Private Function StartPage() As String
  Dim strmpos As Long
  On Error GoTo er
  obj = obj + 1
  location(obj) = Position
  pageNo = pageNo + 1
  pageObj(pageNo) = obj
  
  writepdf (obj & " 0 obj")
  writepdf ("<<")
  writepdf ("/Type /Page")
  writepdf ("/Parent " & Tpages & " 0 R")
  writepdf ("/Resources " & resources & " 0 R")
  obj = obj + 1
  writepdf ("/Contents " & obj & " 0 R")
  writepdf ("/Rotate " & rotate)
  writepdf (">>")
  writepdf ("endobj")
  
  location(obj) = Position
  writepdf (obj & " 0 obj")
  writepdf ("<<")
  writepdf ("/Length " & obj + 1 & " 0 R")
  writepdf (">>")
  writepdf ("stream")
  strmpos = Position
  writepdf ("BT")
  writepdf ("/F1 " & pointSize & " Tf")
  writepdf ("1 0 0 1 50 " & pageHeight - 40 & " Tm")
  writepdf (vertSpace & " TL")
  
  StartPage = strmpos
  Exit Function
er:
  MsgBox Err.Description
End Function

Function endpage(streamstart As Long) As String
  Dim streamEnd As Long
  On Error GoTo er
    writepdf ("ET")
    streamEnd = Position
    writepdf ("endstream")
    writepdf ("endobj")
    obj = obj + 1
    location(obj) = Position
    writepdf (obj & " 0 obj")
    writepdf (streamEnd - streamstart)
    writepdf "endobj"
    lineNo = 0
  Exit Function
er:
  MsgBox Err.Description
End Function

Sub endpdf()
  Dim ty As String, i As Integer, xreF As Long
  On Error GoTo er
    location(root) = Position
    writepdf (root & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Catalog")
    writepdf ("/Pages " & Tpages & " 0 R")
    writepdf (">>")
    writepdf ("endobj")
    location(Tpages) = Position
    writepdf (Tpages & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Pages")
    writepdf ("/Count " & pageNo)
    writepdf ("/MediaBox [ 0 0 " & pageWidth & " " & pageHeight & " ]")
    ty = ("/Kids [ ")
    For i = 1 To pageNo
      ty = ty & pageObj(i) & " 0 R "
    Next i
    ty = ty & "]"
    writepdf (ty)
    writepdf (">>")
    writepdf ("endobj")
    xreF = Position
    writepdf ("0 " & obj + 1)
    writepdf ("0000000000 65535 f ")
    For i = 1 To obj
      writepdf (Format(location(i), "0000000000") & " 00000 n ")
    Next i
    writepdf ("trailer")
    writepdf ("<<")
    writepdf ("/Size " & obj + 1)
    writepdf ("/Root " & root & " 0 R")
    writepdf ("/Info " & info & " 0 R")
    writepdf (">>")
    writepdf ("startxref")
    writepdf (xreF)
    writepdf "%%EOF", True
  Exit Sub
er:
  MsgBox Err.Description
End Sub

Public Function FileExists(ByVal filename As String) As Boolean
  On Error Resume Next
  FileExists = FileLen(filename) > 0
  Err.Clear
End Function

Public Function ReplaceText(Text As String, TextToReplace As String, NewText As String) As String
  Dim mtext As String, SpacePos As Long
  mtext = Text
  SpacePos = InStr(mtext, TextToReplace)
  Do While SpacePos
    mtext = Left(mtext, SpacePos - 1) & NewText & Mid(mtext, SpacePos + Len(TextToReplace))
    SpacePos = InStr(SpacePos + Len(NewText), mtext, TextToReplace)
  Loop
  ReplaceText = mtext
End Function

Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String, DefaultFilename As String) As String
  Dim ofn As OPENFILENAME
  Dim A As Long
  On Local Error Resume Next
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = Form1.hWnd
  ofn.hInstance = App.hInstance
  If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"

  For A = 1 To Len(Filter)
      If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
  Next
  ofn.lpstrFilter = Filter
  ofn.lpstrFile = Space$(254)
  Mid(ofn.lpstrFile, 1, 254) = DefaultFilename
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = InitDir
  ofn.lpstrTitle = Title
  ofn.lpstrDefExt = "pdf"
  ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
  A = GetSaveFileName(ofn)


  If (A) Then
      SaveDialog = Trim$(ofn.lpstrFile)
  Else
      SaveDialog = ""
  End If
End Function

Function OpenDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
  Dim ofn As OPENFILENAME
  Dim A As Long
  On Local Error Resume Next
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = Form1.hWnd
  ofn.hInstance = App.hInstance
  If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"

  For A = 1 To Len(Filter)
      If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
  Next
  ofn.lpstrFilter = Filter
  ofn.lpstrFile = Space$(254)
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = InitDir
  ofn.lpstrTitle = Title
  ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
  A = GetOpenFileName(ofn)

  If (A) Then
      OpenDialog = Trim$(ofn.lpstrFile)
  Else
      OpenDialog = ""
  End If
End Function

