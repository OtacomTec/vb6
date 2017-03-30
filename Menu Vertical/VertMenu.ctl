VERSION 5.00
Begin VB.UserControl VerticalMenu 
   BackColor       =   &H80000010&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2640
   Picture         =   "VertMenu.ctx":0000
   PropertyPages   =   "VertMenu.ctx":0442
   ScaleHeight     =   78
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   176
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   870
      TabIndex        =   1
      Top             =   0
      Width           =   870
   End
   Begin VB.PictureBox picCache 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   990
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   0
      Top             =   315
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image imgDown 
      Height          =   240
      Left            =   2160
      Picture         =   "VertMenu.ctx":046F
      Top             =   540
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Left            =   2160
      Picture         =   "VertMenu.ctx":09B1
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "VerticalMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Dim mMenus As Menus

'Default Property Values:
Const m_def_WhatsThisHelpID = 0
Const m_def_ToolTipText = ""
Const m_def_MousePointer = 0
Const m_def_Enabled = 0
Const m_def_DrawWidth = 0
Const m_def_DrawStyle = 0
Const m_def_DrawMode = 0
Const m_def_CurrentY = 0
Const m_def_CurrentX = 0
Const m_def_BorderStyle = 0
Const m_def_BackStyle = 0
Const m_def_Appearance = 0
Const m_def_AutoRedraw = 0
Const m_def_ClipControls = 0
Const m_def_ScaleWidth = 0
Const m_def_ScaleTop = 0
Const m_def_ScaleMode = 3
Const m_def_ScaleLeft = 0
Const m_def_ScaleHeight = 0
Const m_def_MenusMax = 1
Const m_def_MenuCur = 1
Const m_def_MenuStartup = 1
Const m_def_MenuCaption = "Menu"
Const m_def_MenuItemCaption = "Item"
Const m_def_MenuItemsMax = 1
Const m_def_MenuItemCur = 1

'Property Variables:
Private m_WhatsThisHelpID As Long
Private m_ToolTipText As String
Private m_MousePointer As Integer
Private m_Enabled As Boolean
Private m_DrawWidth As Integer
Private m_DrawStyle As Integer
Private m_DrawMode As Integer
Private m_CurrentY As Single
Private m_CurrentX As Single
Private m_BorderStyle As Integer
Private m_BackStyle As Integer
Private m_ActiveControl As Control
Private m_Appearance As Integer
Private m_AutoRedraw As Boolean
Private m_ClipControls As Boolean
Private m_ScaleWidth As Single
Private m_ScaleTop As Single
Private m_ScaleMode As Integer
Private m_ScaleLeft As Single
Private m_ScaleHeight As Single

Private mlMenusMax As Long
Private mlMenuCur As Long
Private mlMenuStartup As Long
Private msMenuCaption As String
Private msMenuItemCaption As String
Private mpicMenuItemIcon As Picture
Private mlMenuItemsMax As Long
Private mlMenuItemCur As Long
Private mbInitializing As Boolean
Private mbAsyncReadComplete As Boolean
Private mbVBEnvironment As Boolean

' Constants
Const HIT_TYPE_MENU_BUTTON = 1
Const HIT_TYPE_MENUITEM = 2
Const HIT_TYPE_UP_ARROW = 3
Const HIT_TYPE_DOWN_ARROW = 4
Const BUTTON_HEIGHT = 18
Const MOUSE_UP = 1
Const MOUSE_DOWN = -1
Const MOUSE_MOVE = 0
Const MOUSE_IN_CAPTION = -2
Const ICON_SIZE = 32

'Event Declarations:
Event Show()
Event Resize()
Event Hide()
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Paint()
Event MenuItemClick(MenuNumber As Long, MenuItem As Long)

Private Sub picCache_Resize()
    DrawCacheMenuButton
End Sub

' if picMenu considers a second mousedown event as a dblclick, the
' MouseDown event does not file so we need to do it instead
Private Sub picMenu_DblClick()
    Dim POINTAPI As POINTAPI
    Dim lResCod As Long
    
    On Error Resume Next
    lResCod = GetCursorPos(POINTAPI)
    lResCod = ScreenToClient(picMenu.hWnd, POINTAPI)
    picMenu_MouseDown vbLeftButton, 0, CSng(POINTAPI.x), CSng(POINTAPI.y)
End Sub

Private Sub picMenu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lIndex As Long
    Dim lHitType As Long    ' return variable
    
    On Error Resume Next

    If Button = vbLeftButton Then
        With mMenus
            ' currently we only care about MenuButton hits
            ' all others are already processed
            lIndex = .MouseProcess(MOUSE_DOWN, CLng(x), CLng(y), lHitType)
            If lHitType = HIT_TYPE_MENU_BUTTON And lIndex > 0 Then
                MenuCur = lIndex
            End If
        End With
    End If
End Sub

Private Sub picMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ' we don't care about the HitType (an optional parameter)
    mMenus.MouseProcess MOUSE_MOVE, CLng(x), CLng(y)
End Sub

Private Sub picMenu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lMenuItem As Long
    Dim lHitType As Long
    
    On Error Resume Next
    If Button = vbLeftButton Then
        lMenuItem = mMenus.MouseProcess(MOUSE_UP, CLng(x), CLng(y), lHitType)
        If lHitType = HIT_TYPE_MENUITEM And lMenuItem > 0 Then
            picMenu_MouseMove Button, Shift, x, y
            RaiseEvent MenuItemClick(mlMenuCur, lMenuItem)
            picMenu_MouseMove 0, 0, 0, 0
        End If
    End If
End Sub

Private Sub picMenu_Paint()
    On Error Resume Next
    ' using the control with the internet explorer generates a paint
    ' event each time an icon is loaded.  Therefore, don't do the paint
    ' event unless picMenu is visible
    If picMenu.Visible Then
        mMenus.Paint
    End If
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    Dim lSavMenuCur As Long
    Dim lSavMenuItemCur As Long
    On Error Resume Next
    mbAsyncReadComplete = True
    With AsyncProp
        lSavMenuCur = mlMenuCur
        lSavMenuItemCur = mlMenuItemCur
        mlMenuCur = Val(Left$(.PropertyName, 1))
        mlMenuItemCur = Val(Mid$(.PropertyName, 2))
        Set MenuItemIcon = AsyncProp.Value
        mlMenuCur = lSavMenuCur
        mlMenuItemCur = lSavMenuItemCur
    End With
    mbAsyncReadComplete = False
End Sub

Private Sub UserControl_Paint()
    On Error Resume Next
    If Not mbInitializing Then
        picMenu_Paint
    End If
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    Set mMenus = New Menus
    Set mMenus.Menu = picMenu
    Set mMenus.Cache = picCache
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.ScaleMode = vbPixels
    With picMenu
        .ScaleMode = vbPixels
        .Left = 0
        .Top = 0
        .Width = UserControl.ScaleWidth
        .Height = UserControl.ScaleHeight
    End With
    
    With picCache
        .ScaleMode = vbPixels
        .Width = picMenu.Width
        .Height = (BUTTON_HEIGHT * 2) + 33
    End With
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    Set mMenus = Nothing
End Sub

Public Property Get MenusMax() As Long
    On Error Resume Next
    MenusMax = mlMenusMax
End Property

Public Property Let MenusMax(ByVal New_MenusMax As Long)
    Dim l As Long
    Dim lSavMenuCur As Long
    Dim hWnd As Long
    
    On Error Resume Next
    If New_MenusMax < 0 Or New_MenusMax > 6 Then
        Beep
        MsgBox "MenusMax must be between 0 and 6", vbOKOnly
        Exit Property
    End If
    
    UserControl.ScaleMode = vbPixels
    
    Select Case New_MenusMax
        Case mlMenusMax             ' nothing to do
        Case Is > mlMenusMax        ' add menus
            lSavMenuCur = mlMenuCur
            For mlMenuCur = mlMenusMax + 1 To New_MenusMax
                With mMenus
                    .Add "", mlMenuCur, picMenu
                    MenuCaption = m_def_MenuCaption & CStr(mlMenuCur)
                
                    ' set the up/down bitmaps
                    Set .Item(mlMenuCur).UpBitmap = imgUp.Picture
                    Set .Item(mlMenuCur).DownBitmap = imgDown.Picture
                    Set .Item(mlMenuCur).ImageCache = picCache
                    
                    ' add MenuItems to the menu
                    .Item(mlMenuCur).AddMenuItem m_def_MenuItemCaption, 1, mpicMenuItemIcon
                End With
            Next
            mlMenuCur = lSavMenuCur
        Case Is < mlMenusMax        ' delete menus
            For l = mlMenusMax To New_MenusMax + 1 Step -1
                With mMenus
                    .Delete l
                    If New_MenusMax < mlMenuCur Then
                        MenuCur = New_MenusMax
                    End If
                End With
            Next
    End Select
    
    mlMenusMax = New_MenusMax
    mMenus.NumberOfMenusChanged = True
    SetupCache
    UserControl_Paint
    PropertyChanged "MenusMax"
End Property

Public Property Get MenuCur() As Long
    MenuCur = mlMenuCur
End Property

Public Property Let MenuCur(ByVal New_MenuCur As Long)
    On Error Resume Next
    
    ' if we are calling from AsyncReadComplete event, get out of here!
    If mbAsyncReadComplete Then
        Exit Property
    End If
    
    mlMenuCur = New_MenuCur
    mlMenuItemCur = 1           ' reset the menuitem
    With mMenus
        .MenuCur = mlMenuCur
        mlMenuItemsMax = .Item(mlMenuCur).MenuItemCount
        MenuCaption = .Item(mlMenuCur).Caption
    End With
    PropertyChanged "MenuCur"
End Property

Public Property Get MenuStartup() As Long
    On Error Resume Next
    MenuStartup = mlMenuStartup
End Property

Public Property Let MenuStartup(ByVal New_MenuStartup As Long)
    On Error Resume Next
    mlMenuStartup = New_MenuStartup
    PropertyChanged "MenuStartup"
End Property

Public Property Get MenuCaption() As String
    On Error Resume Next
    MenuCaption = msMenuCaption
End Property

Public Property Let MenuCaption(ByVal New_MenuCaption As String)
    On Error Resume Next
    msMenuCaption = New_MenuCaption
    mMenus.Item(mlMenuCur).Caption = New_MenuCaption
    UserControl_Paint
    PropertyChanged "MenuCaption"
End Property

Public Property Get MenuItemCaption() As String
    On Error Resume Next
    msMenuItemCaption = mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Caption
    MenuItemCaption = msMenuItemCaption
End Property

Public Property Let MenuItemCaption(ByVal New_MenuItemCaption As String)
    On Error Resume Next
    With mMenus.Item(mlMenuCur)
        .MenuItemItem(mlMenuItemCur).Caption = New_MenuItemCaption
        msMenuItemCaption = New_MenuItemCaption
    End With
    If Not mbInitializing Then
        picMenu.Cls
        UserControl_Paint
    End If
    PropertyChanged "MenuItemCaption"
End Property

Public Property Get MenuItemIcon() As Picture
    On Error Resume Next
    Set MenuItemIcon = mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Button
End Property

Public Property Set MenuItemIcon(ByVal New_MenuItemIcon As Picture)
    On Error Resume Next
    Set mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Button = New_MenuItemIcon
    If Not mbInitializing Then
        SetupCache
        UserControl_Paint
    End If
    PropertyChanged "MenuItemIcon"
End Property

Public Property Get MenuItemPictureURL() As String
    On Error Resume Next
    MenuItemPictureURL = mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).PictureURL
End Property

Public Property Let MenuItemPictureURL(ByVal New_MenuItemPictureURL As String)
    On Error Resume Next
    mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).PictureURL = New_MenuItemPictureURL
    UserControl.AsyncRead New_MenuItemPictureURL, vbAsyncTypePicture, CStr(mlMenuCur) & CStr(mlMenuItemCur)
    If Err.Number <> 0 Then
    '    Set MenuItemIcon = mpicMenuItemIcon
        Err.Clear
    End If
    PropertyChanged "MenuItemPictureURL"
End Property

Public Property Get MenuItemKey() As String
    On Error Resume Next
    MenuItemKey = mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Key
End Property

Public Property Let MenuItemKey(ByVal New_MenuItemKey As String)
    On Error Resume Next
    mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Key = New_MenuItemKey
    PropertyChanged "MenuItemKey"
End Property

Public Property Get MenuItemTag() As String
    On Error Resume Next
    MenuItemTag = mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Tag
End Property

Public Property Let MenuItemTag(ByVal New_MenuItemTag As String)
    On Error Resume Next
    mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Tag = New_MenuItemTag
    PropertyChanged "MenuItemTag"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Dim l As Long
    
    On Error Resume Next
    
    mbInitializing = True
    mbVBEnvironment = IsThisVB
    
    mMenus.ButtonHeight = BUTTON_HEIGHT             ' set button height for icons
    
    ' set property defaults
    m_Enabled = m_def_Enabled
    m_Appearance = m_def_Appearance
    m_ScaleWidth = m_def_ScaleWidth
    m_ScaleTop = m_def_ScaleTop
    m_ScaleMode = m_def_ScaleMode
    m_ScaleLeft = m_def_ScaleLeft
    m_ScaleHeight = m_def_ScaleHeight
    m_ToolTipText = m_def_ToolTipText
    m_WhatsThisHelpID = m_def_WhatsThisHelpID
    msMenuCaption = m_def_MenuCaption
    msMenuItemCaption = m_def_MenuItemCaption
    mlMenuItemCur = m_def_MenuItemCur
    mlMenuItemsMax = m_def_MenuItemsMax
    
    ProcessDefaultIcon
    
    ' setup the image cache
    With picCache
        .Width = picMenu.Width
        .Height = (BUTTON_HEIGHT * 2) + 33
        .BackColor = BACKGROUND_COLOR
    End With
    picMenu.BackColor = BACKGROUND_COLOR
    
    ' setup the control
    MenusMax = m_def_MenusMax
    MenuCur = m_def_MenuStartup
    MenuStartup = m_def_MenuStartup
    m_WhatsThisHelpID = m_def_WhatsThisHelpID
    m_ToolTipText = m_def_ToolTipText
    m_MousePointer = m_def_MousePointer
    m_Enabled = m_def_Enabled
    m_AutoRedraw = m_def_AutoRedraw
    m_ClipControls = m_def_ClipControls
    
    ' setup the menu caption button and menu item icon cache
    SetupCache

    mbInitializing = False
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim lSavMenuItemCur As Long
    
    On Error Resume Next
    mbInitializing = True
    mbVBEnvironment = IsThisVB
    picMenu.BackColor = BACKGROUND_COLOR
    
    With PropBag
        m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
        m_ToolTipText = .ReadProperty("ToolTipText", m_def_ToolTipText)
        m_WhatsThisHelpID = .ReadProperty("WhatsThisHelpID", m_def_WhatsThisHelpID)
        mlMenuItemCur = m_def_MenuItemCur
        mlMenuItemsMax = m_def_MenuItemsMax
    
        Set mpicMenuItemIcon = .ReadProperty("MenuItemIcon0", Nothing)
        ProcessDefaultIcon
        
        ' setup the image cache
        With picCache
            .Width = UserControl.Width
            .Height = (BUTTON_HEIGHT * 2) + 33
            .BackColor = BACKGROUND_COLOR
        End With
    
        ' add the first menu (which already exists on the form) to the collection
        ' note that calling MenusMax only add and deletes menus other that the 1 item
        ' in the collection
        mMenus.ButtonHeight = BUTTON_HEIGHT
        MenusMax = .ReadProperty("MenusMax", m_def_MenusMax)
        
        ' setup the control arrays
        For mlMenuCur = 1 To mlMenusMax
            MenuCur = mlMenuCur
            msMenuCaption = .ReadProperty("MenuCaption" & CStr(mlMenuCur), m_def_MenuCaption)
            MenuCaption = msMenuCaption
            
            MenuItemsMax = .ReadProperty("MenuItemsMax" & CStr(mlMenuCur), m_def_MenuItemsMax)
            
            lSavMenuItemCur = mlMenuItemCur
            For mlMenuItemCur = 1 To mMenus.Item(mlMenuCur).MenuItemCount
                If mbVBEnvironment Then
                    Set MenuItemIcon = .ReadProperty("MenuItemIcon" & CStr(mlMenuCur) & CStr(mlMenuItemCur), mpicMenuItemIcon)
                Else
                    MenuItemPictureURL = .ReadProperty("MenuItemPictureURL" & CStr(mlMenuCur) & CStr(mlMenuItemCur), "")
                End If
                MenuItemCaption = .ReadProperty("MenuItemCaption" & CStr(mlMenuCur) & CStr(mlMenuItemCur), m_def_MenuItemCaption)
                MenuItemKey = .ReadProperty("MenuItemKey" & CStr(mlMenuCur) & CStr(mlMenuItemCur), "")
                MenuItemTag = .ReadProperty("MenuItemTag" & CStr(mlMenuCur) & CStr(mlMenuItemCur), "")
            Next
            mlMenuItemCur = lSavMenuItemCur
        Next
        ' reset mlMenuCur right away so we don't have errors!
        mlMenuCur = .ReadProperty("MenuCur", m_def_MenuCur)
        
        MenuItemCur = m_def_MenuItemCur
        mlMenuStartup = .ReadProperty("MenuStartup", m_def_MenuStartup)
        MenuStartup = mlMenuStartup
        MenuCur = mlMenuStartup
        m_WhatsThisHelpID = .ReadProperty("WhatsThisHelpID", m_def_WhatsThisHelpID)
        m_ToolTipText = .ReadProperty("ToolTipText", m_def_ToolTipText)
        m_MousePointer = .ReadProperty("MousePointer", m_def_MousePointer)
        m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
        m_AutoRedraw = .ReadProperty("AutoRedraw", m_def_AutoRedraw)
        m_ClipControls = .ReadProperty("ClipControls", m_def_ClipControls)
    End With
    
    ' setup the menu caption button and menu item icon cache
    SetupCache
    
    mbInitializing = False
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim lSavMenuCur As Long
    Dim lSavMenuItemCur As Long
    
    On Error Resume Next
    
    With PropBag
        Call .WriteProperty("Enabled", m_Enabled, m_def_Enabled)
        Call .WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
        Call .WriteProperty("WhatsThisHelpID", m_WhatsThisHelpID, m_def_WhatsThisHelpID)
        Call .WriteProperty("MenusMax", mlMenusMax, m_def_MenusMax)
        Call .WriteProperty("MenuCur", mlMenuCur, m_def_MenuCur)
        Call .WriteProperty("MenuStartup", mlMenuStartup, m_def_MenuStartup)
        
        lSavMenuCur = mlMenuCur
        For mlMenuCur = 1 To mlMenusMax
            Call .WriteProperty("MenuCaption" & CStr(mlMenuCur), mMenus.Item(mlMenuCur).Caption, m_def_MenuCaption)
        
            ' image stuff here
            Call .WriteProperty("MenuItemsMax" & CStr(mlMenuCur), mMenus.Item(mlMenuCur).MenuItemCount, m_def_MenuItemsMax)
            lSavMenuItemCur = mlMenuItemCur
            For mlMenuItemCur = 1 To mMenus.Item(mlMenuCur).MenuItemCount
                If mbVBEnvironment Then
                    Call .WriteProperty("MenuItemIcon" & CStr(mlMenuCur) & CStr(mlMenuItemCur), MenuItemIcon, Nothing)
                Else
                    Call .WriteProperty("MenuItemPictureURL" & CStr(mlMenuCur) & CStr(mlMenuItemCur), MenuItemPictureURL, "")
                End If
                Call .WriteProperty("MenuItemCaption" & CStr(mlMenuCur) & CStr(mlMenuItemCur), MenuItemCaption, m_def_MenuItemCaption)
                Call .WriteProperty("MenuItemKey" & CStr(mlMenuCur) & CStr(mlMenuItemCur), MenuItemKey, "")
                Call .WriteProperty("MenuItemTag" & CStr(mlMenuCur) & CStr(mlMenuItemCur), MenuItemTag, "")
            Next
            mlMenuItemCur = lSavMenuItemCur
        Next
        mlMenuCur = lSavMenuCur
        Call .WriteProperty("MenuItemIcon0", mpicMenuItemIcon, mpicMenuItemIcon)
        Call .WriteProperty("WhatsThisHelpID", m_WhatsThisHelpID, m_def_WhatsThisHelpID)
        Call .WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
        Call .WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
        Call .WriteProperty("Enabled", m_Enabled, m_def_Enabled)
        Call .WriteProperty("AutoRedraw", m_AutoRedraw, m_def_AutoRedraw)
        Call .WriteProperty("ClipControls", m_ClipControls, m_def_ClipControls)
    End With
End Sub

Public Property Get MenuItemsMax() As Long
    On Error Resume Next
    MenuItemsMax = mlMenuItemsMax
End Property

Public Property Let MenuItemsMax(ByVal New_MenuItemsMax As Long)
    Dim l As Long
    Dim lSavMenuItemCur As Long
    
    On Error Resume Next
    If New_MenuItemsMax < 0 Or New_MenuItemsMax > 10 Then
        Beep
        MsgBox "MenuItemsMax must be between 0 and 10", vbOKOnly
        Exit Property
    End If
    
    lSavMenuItemCur = mlMenuItemCur
    Select Case New_MenuItemsMax
        Case mlMenuItemsMax             ' nothing to do
        Case Is > mlMenuItemsMax        ' add menus
            With mMenus.Item(mlMenuCur)
                For mlMenuItemCur = mlMenuItemsMax + 1 To New_MenuItemsMax
                    .AddMenuItem m_def_MenuItemCaption, mlMenuItemCur, mpicMenuItemIcon
                    MenuItemCaption = m_def_MenuItemCaption & CStr(mlMenuItemCur)
                Next
                mlMenuItemCur = lSavMenuItemCur
            End With
        Case Is < mlMenuItemsMax        ' delete menus
            With mMenus.Item(mlMenuCur)
                For mlMenuItemCur = mlMenuItemsMax To New_MenuItemsMax + 1 Step -1
                    .DeleteMenuItem mlMenuItemCur
                Next
                mlMenuItemCur = lSavMenuItemCur
                If New_MenuItemsMax < mlMenuItemCur Then
                    mlMenuItemCur = New_MenuItemsMax
                End If
            End With
    End Select
    ' reset the caption in the properties window
    mlMenuItemsMax = New_MenuItemsMax
    SetupCache
    UserControl_Paint
    PropertyChanged "MenuItemsMax"
End Property

Public Property Get MenuItemCur() As Long
    On Error Resume Next
    MenuItemCur = mlMenuItemCur
End Property

Public Property Let MenuItemCur(ByVal New_MenuItemCur As Long)
    On Error Resume Next
    
    ' if we are calling from AsyncReadComplete event, get out of here!
    If mbAsyncReadComplete Then
        Exit Property
    End If
    
    If New_MenuItemCur > mlMenuItemsMax Then
        Beep
        MsgBox "The current item must be between 0 and MenuItemsMax", vbOKOnly
        Exit Property
    End If
    mlMenuItemCur = New_MenuItemCur
    PropertyChanged "MenuItemCur"
End Property

Public Sub SetupCache()
    Dim lMenuItemCount As Long
    Dim lMIndex As Long
    Dim lMMax As Long
    Dim lMIIndex As Long
    Dim lMIMax As Long
    Dim lIconIndex As Long
    Const I_OFFSET = BUTTON_HEIGHT * 2 + ICON_SIZE

    On Error Resume Next
    
    picCache.Cls
    DrawCacheMenuButton
    
    ' total MenuItems on the control
    lMenuItemCount = mMenus.TotalMenuItems
    
    With picCache
        .ScaleMode = vbPixels
        
        ' set the height for a menu button, space for an unpainted button
        ' space for an unpainted icon and all the MenuItem icons
        .Height = BUTTON_HEIGHT * 2 + (lMenuItemCount + 1) * ICON_SIZE

        ' loop thru the menus getting each icon for each MenuItem
        lMMax = mMenus.Count
        lIconIndex = 0
        For lMIndex = 1 To lMMax
            lMIMax = mMenus.Item(lMIndex).MenuItemCount
            For lMIIndex = 1 To lMIMax
                lIconIndex = lIconIndex + 1
                picCache.PaintPicture mMenus.Item(lMIndex).MenuItemItem(lMIIndex).Button, _
                    0, I_OFFSET + (lIconIndex - 1) * ICON_SIZE, ICON_SIZE, ICON_SIZE, 0, 0
            Next
        Next
    End With
End Sub

Private Sub ProcessDefaultIcon()
    ' UserControl contains the default picture
    ' set it into mpicMenuItemIcon to use as the default icon
    ' (it will be written to the property bag later)
    ' then delete UserControl.Picture
    ' note that if mpicMenuItemIcon is nothing, then we are reading from
    On Error Resume Next
    If mpicMenuItemIcon Is Nothing Then
        Set mpicMenuItemIcon = UserControl.Picture
    End If
    UserControl.Picture = LoadPicture()
End Sub

Private Sub DrawCacheMenuButton()
    Dim RECT As RECT
    
    RECT.Left = 0
    RECT.Top = 0
    RECT.Right = picCache.ScaleWidth
    RECT.Bottom = BUTTON_HEIGHT
    DrawEdge picCache.hdc, RECT, BDR_RAISED, BF_RECT Or BF_MIDDLE
End Sub

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
    WhatsThisHelpID = m_WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
    m_WhatsThisHelpID = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl_Paint
End Sub

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    m_MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get ClipControls() As Boolean
Attribute ClipControls.VB_Description = "Determines whether graphics methods in Paint events repaint an entire object or newly exposed areas."
    ClipControls = m_ClipControls
End Property

Public Property Let ClipControls(ByVal New_ClipControls As Boolean)
    m_ClipControls = New_ClipControls
    PropertyChanged "ClipControls"
End Property

Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
    dlgAbout.Show vbModal
    Unload dlgAbout
    Set dlgAbout = Nothing
End Sub

' we need to if we are running in VB or a browser
' VB supports this extender object while a browser doesn't
' note:  we can't read icons from the property bag using a browser - GPF's
Private Function IsThisVB() As Boolean
    Dim obj As Object

    On Error Resume Next
    Set UserControl.Extender.Parent = obj
    IsThisVB = (Err.Number = 0)
    Set obj = Nothing
    Err.Clear
End Function
