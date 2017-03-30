Attribute VB_Name = "modIcon"
Option Explicit

'Icon Sizes in pixels
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const MAX_PATH = 260

Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type


'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal x&, ByVal y&, ByVal flags&) As Long


'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO

Public Function GetIcon(FileName As String, Index As Long) As Long
    '---------------------------------------------------------------------
    'Extract an individual icon
    '---------------------------------------------------------------------
    Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
    Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
    Dim r As Long
    
    
    'Get a handle to the small icon
    hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
             BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    'Get a handle to the large icon
    hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
             BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    
    'If the handle(s) exists, load it into the picture box(es)
    If hLIcon <> 0 Then
      'Small Icon
      With frmMain.pic16
        Set .Picture = LoadPicture("")
        .AutoRedraw = True
        r = ImageList_Draw(hSIcon, ShInfo.iIcon, frmMain.pic16.hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      
      Set imgObj = frmMain.img16.ListImages.Add(Index, , frmMain.pic16.Image)
    End If
End Function

Private Sub ShowIcons()
    '-----------------------------------------
    'Show the icons in the lvw
    '-----------------------------------------
    On Error Resume Next
    
    Dim Item As ListItem
    With frmMain.lvAttachments
      '.ListItems.Clear
      .SmallIcons = frmMain.img16   'Small
      For Each Item In .ListItems
        Item.Icon = Item.Index
        Item.SmallIcon = Item.Index
      Next
    End With

End Sub
