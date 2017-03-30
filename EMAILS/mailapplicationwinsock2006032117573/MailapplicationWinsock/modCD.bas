Attribute VB_Name = "modCD"
Option Explicit                              'Force variable declaration

                                             '//
                                             '// Structures
                                             '//

Private Type OPENFILENAME
    lStructSize As Long
    hwnd As Long
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
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type COLORSTRUC
    lStructSize As Long
    hwnd As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const LF_FACESIZE = 32               'Declare lf_facesize with value 32 as constant for local use

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type FONTSTRUC
    lStructSize As Long
    hwnd As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFreq As Long
End Type

Private Type PRINTDLGSTRUC
    lStructSize As Long
    hwnd As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Public Type PRINTPROPS
    Cancel As Boolean
    Device As String
    Copies As Integer
    FromPage As Integer
    ToPage As Integer
    ToFile As Boolean
    Range As Integer
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

                                             '//
                                             '// Win32s
                                             '//

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long 'Declare the getopenfilename API for local use
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long 'Declare the getsavefilename API for local use
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGSTRUC) As Long 'Declare the printdlg API for local use
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As COLORSTRUC) As Long 'Declare the choosecolor API for local use
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As FONTSTRUC) As Long 'Declare the choosefont API for local use
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long 'Declare the globalalloc API for local use
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long 'Declare the globalfree API for local use
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long 'Declare the globallock API for local use
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long 'Declare the globalunlock API for local use
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long) 'Declare the copymemory API for local use
Private Declare Function ConnectToPrinterDlg Lib "winspool.drv" (ByVal hwnd As Long, ByVal Flags As Long) As Long 'Declare the connecttoprinterdlg API for local use
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long 'Declare the shgetpathfromidlist API for local use
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long 'Declare the shgetspecialfolderlocation API for local use
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long 'Declare the writeprofilestring API for local use
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long 'Declare the getprofilestring API for local use
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long 'Declare the sendmessagebystring API for local use
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long 'Declare the declare function winhelp API declare function winhelp lib "user32" alias "winhelpa" (byval hwnd as long, byval lphelpfile as string, byval wcommand as long, byval dwdata as long) as lon


                                             '//
                                             '// Constants (Public for Print Properties Structure)
                                             '//

Public Const ppRangeAll = 0                  'Declare pprangeall with value 0 as constant with max scope
Public Const ppRangePages = 1                'Declare pprangepages with value 1 as constant with max scope
Public Const ppRangeSelection = 2            'Declare pprangeselection with value 2 as constant with max scope

                                             '//
                                             '// Constants (Public for Print Dialog Box)
                                             '//

Public Const PD_NOSELECTION = &H4            'Declare pd_noselection with value &h4 as constant with max scope
Public Const PD_DISABLEPRINTTOFILE = &H80000 'Declare pd_disableprinttofile with value &h80000 as constant with max scope
Public Const PD_PRINTTOFILE = &H20           'Declare pd_printtofile with value &h20 as constant with max scope
Public Const PD_RETURNDC = &H100             'Declare pd_returndc with value &h100 as constant with max scope
Public Const PD_RETURNDEFAULT = &H400        'Declare pd_returndefault with value &h400 as constant with max scope
Public Const PD_RETURNIC = &H200             'Declare pd_returnic with value &h200 as constant with max scope
Public Const PD_SELECTION = &H1              'Declare pd_selection with value &h1 as constant with max scope
Public Const PD_SHOWHELP = &H800             'Declare pd_showhelp with value &h800 as constant with max scope
Public Const PD_NOPAGENUMS = &H8             'Declare pd_nopagenums with value &h8 as constant with max scope
Public Const PD_PAGENUMS = &H2               'Declare pd_pagenums with value &h2 as constant with max scope

                                             '//
                                             '// Constants (Public for WinHelp)
                                             '//

Public Const HELP_COMMAND = &H102&           'Declare help_command with value &h102& as constant with max scope
Public Const HELP_CONTENTS = &H3&            'Declare help_contents with value &h3& as constant with max scope
Public Const HELP_CONTEXT = &H1              'Declare help_context with value &h1 as constant with max scope
Public Const HELP_CONTEXTPOPUP = &H8&        'Declare help_contextpopup with value &h8& as constant with max scope
Public Const HELP_FORCEFILE = &H9&           'Declare help_forcefile with value &h9& as constant with max scope
Public Const HELP_HELPONHELP = &H4           'Declare help_helponhelp with value &h4 as constant with max scope
Public Const HELP_INDEX = &H3                'Declare help_index with value &h3 as constant with max scope
Public Const HELP_KEY = &H101                'Declare help_key with value &h101 as constant with max scope
Public Const HELP_MULTIKEY = &H201&          'Declare help_multikey with value &h201& as constant with max scope
Public Const HELP_PARTIALKEY = &H105&        'Declare help_partialkey with value &h105& as constant with max scope
Public Const HELP_QUIT = &H2                 'Declare help_quit with value &h2 as constant with max scope
Public Const HELP_SETCONTENTS = &H5&         'Declare help_setcontents with value &h5& as constant with max scope
Public Const HELP_SETINDEX = &H5             'Declare help_setindex with value &h5 as constant with max scope
Public Const HELP_SETWINPOS = &H203&         'Declare help_setwinpos with value &h203& as constant with max scope


                                             '//
                                             '// Constants (Private)
                                             '//

Private Const FW_BOLD = 700                  'Declare fw_bold with value 700 as constant for local use
Private Const GMEM_MOVEABLE = &H2            'Declare gmem_moveable with value &h2 as constant for local use
Private Const GMEM_ZEROINIT = &H40           'Declare gmem_zeroinit with value &h40 as constant for local use
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT) 'Declare ghnd with value (gmem_moveable or gmem_zeroinit) as constant for local use
Private Const OFN_ALLOWMULTISELECT = &H200   'Declare ofn_allowmultiselect with value &h200 as constant for local use
Private Const OFN_CREATEPROMPT = &H2000      'Declare ofn_createprompt with value &h2000 as constant for local use
Private Const OFN_ENABLEHOOK = &H20          'Declare ofn_enablehook with value &h20 as constant for local use
Private Const OFN_ENABLETEMPLATE = &H40      'Declare ofn_enabletemplate with value &h40 as constant for local use
Private Const OFN_ENABLETEMPLATEHANDLE = &H80 'Declare ofn_enabletemplatehandle with value &h80 as constant for local use
Private Const OFN_EXPLORER = &H80000         'Declare ofn_explorer with value &h80000 as constant for local use
Private Const OFN_EXTENSIONDIFFERENT = &H400 'Declare ofn_extensiondifferent with value &h400 as constant for local use
Private Const OFN_FILEMUSTEXIST = &H1000     'Declare ofn_filemustexist with value &h1000 as constant for local use
Private Const OFN_HIDEREADONLY = &H4         'Declare ofn_hidereadonly with value &h4 as constant for local use
Private Const OFN_LONGNAMES = &H200000       'Declare ofn_longnames with value &h200000 as constant for local use
Private Const OFN_NOCHANGEDIR = &H8          'Declare ofn_nochangedir with value &h8 as constant for local use
Private Const OFN_NODEREFERENCELINKS = &H100000 'Declare ofn_nodereferencelinks with value &h100000 as constant for local use
Private Const OFN_NOLONGNAMES = &H40000      'Declare ofn_nolongnames with value &h40000 as constant for local use
Private Const OFN_NONETWORKBUTTON = &H20000  'Declare ofn_nonetworkbutton with value &h20000 as constant for local use
Private Const OFN_NOREADONLYRETURN = &H8000  'Declare ofn_noreadonlyreturn with value &h8000 as constant for local use
Private Const OFN_NOTESTFILECREATE = &H10000 'Declare ofn_notestfilecreate with value &h10000 as constant for local use
Private Const OFN_NOVALIDATE = &H100         'Declare ofn_novalidate with value &h100 as constant for local use
Private Const OFN_OVERWRITEPROMPT = &H2      'Declare ofn_overwriteprompt with value &h2 as constant for local use
Private Const OFN_PATHMUSTEXIST = &H800      'Declare ofn_pathmustexist with value &h800 as constant for local use
Private Const OFN_READONLY = &H1             'Declare ofn_readonly with value &h1 as constant for local use
Private Const OFN_SHAREAWARE = &H4000        'Declare ofn_shareaware with value &h4000 as constant for local use
Private Const OFN_SHAREFALLTHROUGH = 2       'Declare ofn_sharefallthrough with value 2 as constant for local use
Private Const OFN_SHARENOWARN = 1            'Declare ofn_sharenowarn with value 1 as constant for local use
Private Const OFN_SHAREWARN = 0              'Declare ofn_sharewarn with value 0 as constant for local use
Private Const OFN_SHOWHELP = &H10            'Declare ofn_showhelp with value &h10 as constant for local use
Private Const PD_ALLPAGES = &H0              'Declare pd_allpages with value &h0 as constant for local use
Private Const PD_COLLATE = &H10              'Declare pd_collate with value &h10 as constant for local use
Private Const PD_ENABLEPRINTHOOK = &H1000    'Declare pd_enableprinthook with value &h1000 as constant for local use
Private Const PD_ENABLEPRINTTEMPLATE = &H4000 'Declare pd_enableprinttemplate with value &h4000 as constant for local use
Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000 'Declare pd_enableprinttemplatehandle with value &h10000 as constant for local use
Private Const PD_ENABLESETUPHOOK = &H2000    'Declare pd_enablesetuphook with value &h2000 as constant for local use
Private Const PD_ENABLESETUPTEMPLATE = &H8000 'Declare pd_enablesetuptemplate with value &h8000 as constant for local use
Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000 'Declare pd_enablesetuptemplatehandle with value &h20000 as constant for local use
Private Const PD_HIDEPRINTTOFILE = &H100000  'Declare pd_hideprinttofile with value &h100000 as constant for local use
Private Const PD_NONETWORKBUTTON = &H200000  'Declare pd_nonetworkbutton with value &h200000 as constant for local use
Private Const PD_PRINTSETUP = &H40           'Declare pd_printsetup with value &h40 as constant for local use
Private Const PD_USEDEVMODECOPIES = &H40000  'Declare pd_usedevmodecopies with value &h40000 as constant for local use
Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000 'Declare pd_usedevmodecopiesandcollate with value &h40000 as constant for local use
Private Const PD_NOWARNING = &H80            'Declare pd_nowarning with value &h80 as constant for local use
Private Const CF_ANSIONLY = &H400&           'Declare cf_ansionly with value &h400& as constant for local use
Private Const CF_APPLY = &H200&              'Declare cf_apply with value &h200& as constant for local use
Private Const CF_BITMAP = 2                  'Declare cf_bitmap with value 2 as constant for local use
Private Const CF_PRINTERFONTS = &H2          'Declare cf_printerfonts with value &h2 as constant for local use
Private Const CF_PRIVATEFIRST = &H200        'Declare cf_privatefirst with value &h200 as constant for local use
Private Const CF_PRIVATELAST = &H2FF         'Declare cf_privatelast with value &h2ff as constant for local use
Private Const CF_RIFF = 11                   'Declare cf_riff with value 11 as constant for local use
Private Const CF_SCALABLEONLY = &H20000      'Declare cf_scalableonly with value &h20000 as constant for local use
Private Const CF_SCREENFONTS = &H1           'Declare cf_screenfonts with value &h1 as constant for local use
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS) 'Declare cf_both with value (cf_screenfonts or cf_printerfonts) as constant for local use
Private Const CF_DIB = 8                     'Declare cf_dib with value 8 as constant for local use
Private Const CF_DIF = 5                     'Declare cf_dif with value 5 as constant for local use
Private Const CF_DSPBITMAP = &H82            'Declare cf_dspbitmap with value &h82 as constant for local use
Private Const CF_DSPENHMETAFILE = &H8E       'Declare cf_dspenhmetafile with value &h8e as constant for local use
Private Const CF_DSPMETAFILEPICT = &H83      'Declare cf_dspmetafilepict with value &h83 as constant for local use
Private Const CF_DSPTEXT = &H81              'Declare cf_dsptext with value &h81 as constant for local use
Private Const CF_EFFECTS = &H100&            'Declare cf_effects with value &h100& as constant for local use
Private Const CF_ENABLEHOOK = &H8&           'Declare cf_enablehook with value &h8& as constant for local use
Private Const CF_ENABLETEMPLATE = &H10&      'Declare cf_enabletemplate with value &h10& as constant for local use
Private Const CF_ENABLETEMPLATEHANDLE = &H20& 'Declare cf_enabletemplatehandle with value &h20& as constant for local use
Private Const CF_ENHMETAFILE = 14            'Declare cf_enhmetafile with value 14 as constant for local use
Private Const CF_FIXEDPITCHONLY = &H4000&    'Declare cf_fixedpitchonly with value &h4000& as constant for local use
Private Const CF_FORCEFONTEXIST = &H10000    'Declare cf_forcefontexist with value &h10000 as constant for local use
Private Const CF_GDIOBJFIRST = &H300         'Declare cf_gdiobjfirst with value &h300 as constant for local use
Private Const CF_GDIOBJLAST = &H3FF          'Declare cf_gdiobjlast with value &h3ff as constant for local use
Private Const CF_INITTOLOGFONTSTRUCT = &H40& 'Declare cf_inittologfontstruct with value &h40& as constant for local use
Private Const CF_LIMITSIZE = &H2000&         'Declare cf_limitsize with value &h2000& as constant for local use
Private Const CF_METAFILEPICT = 3            'Declare cf_metafilepict with value 3 as constant for local use
Private Const CF_NOFACESEL = &H80000         'Declare cf_nofacesel with value &h80000 as constant for local use
Private Const CF_NOVERTFONTS = &H1000000     'Declare cf_novertfonts with value &h1000000 as constant for local use
Private Const CF_NOVECTORFONTS = &H800&      'Declare cf_novectorfonts with value &h800& as constant for local use
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS 'Declare cf_nooemfonts with value cf_novectorfonts as constant for local use
Private Const CF_NOSCRIPTSEL = &H800000      'Declare cf_noscriptsel with value &h800000 as constant for local use
Private Const CF_NOSIMULATIONS = &H1000&     'Declare cf_nosimulations with value &h1000& as constant for local use
Private Const CF_NOSIZESEL = &H200000        'Declare cf_nosizesel with value &h200000 as constant for local use
Private Const CF_NOSTYLESEL = &H100000       'Declare cf_nostylesel with value &h100000 as constant for local use
Private Const CF_OEMTEXT = 7                 'Declare cf_oemtext with value 7 as constant for local use
Private Const CF_OWNERDISPLAY = &H80         'Declare cf_ownerdisplay with value &h80 as constant for local use
Private Const CF_PALETTE = 9                 'Declare cf_palette with value 9 as constant for local use
Private Const CF_PENDATA = 10                'Declare cf_pendata with value 10 as constant for local use
Private Const CF_SCRIPTSONLY = CF_ANSIONLY   'Declare cf_scriptsonly with value cf_ansionly as constant for local use
Private Const CF_SELECTSCRIPT = &H400000     'Declare cf_selectscript with value &h400000 as constant for local use
Private Const CF_SHOWHELP = &H4&             'Declare cf_showhelp with value &h4& as constant for local use
Private Const CF_SYLK = 4                    'Declare cf_sylk with value 4 as constant for local use
Private Const CF_TEXT = 1                    'Declare cf_text with value 1 as constant for local use
Private Const CF_TIFF = 6                    'Declare cf_tiff with value 6 as constant for local use
Private Const CF_TTONLY = &H40000            'Declare cf_ttonly with value &h40000 as constant for local use
Private Const CF_UNICODETEXT = 13            'Declare cf_unicodetext with value 13 as constant for local use
Private Const CF_USESTYLE = &H80&            'Declare cf_usestyle with value &h80& as constant for local use
Private Const CF_WAVE = 12                   'Declare cf_wave with value 12 as constant for local use
Private Const CF_WYSIWYG = &H8000            'Declare cf_wysiwyg with value &h8000 as constant for local use
Private Const CFERR_CHOOSEFONTCODES = &H2000 'Declare cferr_choosefontcodes with value &h2000 as constant for local use
Private Const CFERR_MAXLESSTHANMIN = &H2002  'Declare cferr_maxlessthanmin with value &h2002 as constant for local use
Private Const CFERR_NOFONTS = &H2001         'Declare cferr_nofonts with value &h2001 as constant for local use
Private Const CC_ANYCOLOR = &H100            'Declare cc_anycolor with value &h100 as constant for local use
Private Const CC_CHORD = 4                   'Declare cc_chord with value 4 as constant for local use
Private Const CC_CIRCLES = 1                 'Declare cc_circles with value 1 as constant for local use
Private Const CC_ELLIPSES = 8                'Declare cc_ellipses with value 8 as constant for local use
Private Const CC_ENABLEHOOK = &H10           'Declare cc_enablehook with value &h10 as constant for local use
Private Const CC_ENABLETEMPLATE = &H20       'Declare cc_enabletemplate with value &h20 as constant for local use
Private Const CC_ENABLETEMPLATEHANDLE = &H40 'Declare cc_enabletemplatehandle with value &h40 as constant for local use
Private Const CC_FULLOPEN = &H2              'Declare cc_fullopen with value &h2 as constant for local use
Private Const CC_INTERIORS = 128             'Declare cc_interiors with value 128 as constant for local use
Private Const CC_NONE = 0                    'Declare cc_none with value 0 as constant for local use
Private Const CC_PIE = 2                     'Declare cc_pie with value 2 as constant for local use
Private Const CC_PREVENTFULLOPEN = &H4       'Declare cc_preventfullopen with value &h4 as constant for local use
Private Const CC_RGBINIT = &H1               'Declare cc_rgbinit with value &h1 as constant for local use
Private Const CC_ROUNDRECT = 256
Private Const CC_SHOWHELP = &H8              'Declare cc_showhelp with value &h8 as constant for local use
Private Const CC_SOLIDCOLOR = &H80           'Declare cc_solidcolor with value &h80 as constant for local use
Private Const CC_STYLED = 32                 'Declare cc_styled with value 32 as constant for local use
Private Const CC_WIDE = 16                   'Declare cc_wide with value 16 as constant for local use
Private Const CC_WIDESTYLED = 64             'Declare cc_widestyled with value 64 as constant for local use
Private Const CCERR_CHOOSECOLORCODES = &H5000 'Declare ccerr_choosecolorcodes with value &h5000 as constant for local use
Private Const LOGPIXELSY = 90                'Declare logpixelsy with value 90 as constant for local use
Private Const CCHDEVICENAME = 32             'Declare cchdevicename with value 32 as constant for local use
Private Const CCHFORMNAME = 32               'Declare cchformname with value 32 as constant for local use
Private Const SIMULATED_FONTTYPE = &H8000    'Declare simulated_fonttype with value &h8000 as constant for local use
Private Const PRINTER_FONTTYPE = &H4000      'Declare printer_fonttype with value &h4000 as constant for local use
Private Const SCREEN_FONTTYPE = &H2000       'Declare screen_fonttype with value &h2000 as constant for local use
Private Const BOLD_FONTTYPE = &H100          'Declare bold_fonttype with value &h100 as constant for local use
Private Const ITALIC_FONTTYPE = &H200        'Declare italic_fonttype with value &h200 as constant for local use
Private Const REGULAR_FONTTYPE = &H400       'Declare regular_fonttype with value &h400 as constant for local use
Private Const WM_CHOOSEFONT_GETLOGFONT = (&H400 + 1) 'Declare wm_choosefont_getlogfont with value (&h400 + 1) as constant for local use
Private Const LBSELCHSTRING = "commdlg_LBSelChangedNotify" 'Declare lbselchstring with value "commdlg_lbselchangednotify" as constant for local use
Private Const SHAREVISTRING = "commdlg_ShareViolation" 'Declare sharevistring with value "commdlg_shareviolation" as constant for local use
Private Const FILEOKSTRING = "commdlg_FileNameOK" 'Declare fileokstring with value "commdlg_filenameok" as constant for local use
Private Const COLOROKSTRING = "commdlg_ColorOK" 'Declare colorokstring with value "commdlg_colorok" as constant for local use
Private Const SETRGBSTRING = "commdlg_SetRGBColor" 'Declare setrgbstring with value "commdlg_setrgbcolor" as constant for local use
Private Const FINDMSGSTRING = "commdlg_FindReplace" 'Declare findmsgstring with value "commdlg_findreplace" as constant for local use
Private Const HELPMSGSTRING = "commdlg_help" 'Declare helpmsgstring with value "commdlg_help" as constant for local use
Private Const CD_LBSELNOITEMS = -1           'Declare cd_lbselnoitems with value -1 as constant for local use
Private Const CD_LBSELCHANGE = 0             'Declare cd_lbselchange with value 0 as constant for local use
Private Const CD_LBSELSUB = 1                'Declare cd_lbselsub with value 1 as constant for local use
Private Const CD_LBSELADD = 2                'Declare cd_lbseladd with value 2 as constant for local use
Private Const NOERROR = 0                    'Declare noerror with value 0 as constant for local use
Private Const CSIDL_DESKTOP = &H0            'Declare csidl_desktop with value &h0 as constant for local use
Private Const CSIDL_PROGRAMS = &H2           'Declare csidl_programs with value &h2 as constant for local use
Private Const CSIDL_CONTROLS = &H3           'Declare csidl_controls with value &h3 as constant for local use
Private Const CSIDL_PRINTERS = &H4           'Declare csidl_printers with value &h4 as constant for local use
Private Const CSIDL_PERSONAL = &H5           'Declare csidl_personal with value &h5 as constant for local use
Private Const CSIDL_FAVORITES = &H6          'Declare csidl_favorites with value &h6 as constant for local use
Private Const CSIDL_STARTUP = &H7            'Declare csidl_startup with value &h7 as constant for local use
Private Const CSIDL_RECENT = &H8             'Declare csidl_recent with value &h8 as constant for local use
Private Const CSIDL_SENDTO = &H9             'Declare csidl_sendto with value &h9 as constant for local use
Private Const CSIDL_BITBUCKET = &HA          'Declare csidl_bitbucket with value &ha as constant for local use
Private Const CSIDL_STARTMENU = &HB          'Declare csidl_startmenu with value &hb as constant for local use
Private Const CSIDL_DESKTOPDIRECTORY = &H10  'Declare csidl_desktopdirectory with value &h10 as constant for local use
Private Const CSIDL_DRIVES = &H11            'Declare csidl_drives with value &h11 as constant for local use
Private Const CSIDL_NETWORK = &H12           'Declare csidl_network with value &h12 as constant for local use
Private Const CSIDL_NETHOOD = &H13           'Declare csidl_nethood with value &h13 as constant for local use
Private Const CSIDL_FONTS = &H14             'Declare csidl_fonts with value &h14 as constant for local use
Private Const CSIDL_TEMPLATES = &H15         'Declare csidl_templates with value &h15 as constant for local use
Private Const HWND_BROADCAST = &HFFFF&       'Declare hwnd_broadcast with value &hffff& as constant for local use
Private Const WM_WININICHANGE = &H1A         'Declare wm_wininichange with value &h1a as constant for local use

                                             'alle bladeropties
Private Const BIF_RETURNONLYFSDIRS = &H1     'Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed.
Private Const BIF_DONTGOBELOWDOMAIN = &H2    'Do not include network folders below the domain level in the tree view control.
Private Const BIF_STATUSTEXT = &H4           'Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box.
Private Const BIF_RETURNFSANCESTORS As Long = &H8 'Only return file system ancestors. If the user selects anything other than a file system ancestor, the OK button is grayed.
Private Const BIF_EDITBOX As Long = &H10     'Version 4.71. The browse dialog includes an edit control in which the user can type the name of an item.
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000 'Only return computers. If the user selects anything other than a computer, the OK button is grayed.
Private Const BIF_BROWSEFORPRINTER As Long = &H2000 'Only return printers. If the user selects anything other than a printer, the OK button is grayed.
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000 'The browse dialog will display files as well as folders.


Public Sub SetDefaultPrinter(objPrn As Printer)

    Dim X As Long, szTmp As String           'Declare x for local use as long, sztmp as string

    szTmp = objPrn.DeviceName & "," & objPrn.DriverName & "," & objPrn.Port
    X = WriteProfileString("windows", "device", szTmp)
    X = SendMessageByString(HWND_BROADCAST, WM_WININICHANGE, 0&, "windows")

End Sub
Public Function GetDefaultPrinter() As String

    Dim X As Long, szTmp As String, dwBuf As Long 'Declare x for local use as long, sztmp as string, dwbuf as long

    dwBuf = 1024
    szTmp = Space(dwBuf + 1)
    X = GetProfileString("windows", "device", "", szTmp, dwBuf)
    GetDefaultPrinter = Trim(Left(szTmp, X))

End Function
Public Sub ResetDefaultPrinter(szBuf As String)

    Dim X As Long                            'Declare x for local use as long

    X = WriteProfileString("windows", "device", szBuf)
    X = SendMessageByString(HWND_BROADCAST, WM_WININICHANGE, 0&, "windows")

End Sub
Public Function BrowseFolder(f As Form, szDialogTitle As String, optie As String) As String

    Dim X As Long, BI As BROWSEINFO, dwIList As Long, szPath As String, wPos As Integer 'Declare x for local use as long, bi as browseinfo, dwilist as long, szpath as string, wpos as integer

    BI.hOwner = f.hwnd
    BI.lpszTitle = szDialogTitle
    BI.ulFlags = optie
    dwIList = SHBrowseForFolder(BI)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    If X Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left$(szPath, wPos - 1)
    Else
        BrowseFolder = ""                    'Empty browsefolder
    End If

End Function
Public Function DialogConnectToPrinter(f As Form) As Boolean

    Dim X As Long                            'Declare x for local use as long
    DialogConnectToPrinter = True
    X = ConnectToPrinterDlg(f.hwnd, 0)

End Function
Private Function ByteToString(aBytes() As Byte) As String

    Dim dwBytePoint As Long, dwByteVal As Long, szOut As String 'Declare dwbytepoint for local use as long, dwbyteval as long, szout as string

    dwBytePoint = LBound(aBytes)

    While dwBytePoint <= UBound(aBytes)

        dwByteVal = aBytes(dwBytePoint)

        If dwByteVal = 0 Then
            ByteToString = szOut
            Exit Function                    'Leave this function
        Else
            szOut = szOut & Chr$(dwByteVal)  'Add chr$(dwbyteval) to szout
        End If

        dwBytePoint = dwBytePoint + 1        'Add 1 to dwbytepoint

    Wend

    ByteToString = szOut

End Function
Public Function DialogColor(f As Form, c As Control) As Boolean

    Dim X As Long, CS As COLORSTRUC, CustColor(16) As Long 'Declare x for local use as long, cs as colorstruc, custcolor(16) as long

    CS.lStructSize = Len(CS)
    CS.hwnd = f.hwnd
    CS.hInstance = App.hInstance
    CS.Flags = CC_SOLIDCOLOR
    CS.lpCustColors = String$(16 * 4, 0)
    X = ChooseColor(CS)
    If X = 0 Then
        DialogColor = False
    Else
        DialogColor = True
        c.BackColor = CS.rgbResult
    End If

End Function


Public Function DialogFile(f As Form, wMode As Integer, szDialogTitle As String, szFilename As String, szFilter As String, szDefDir As String, szDefExt As String) As String

    Dim X As Long, OFN As OPENFILENAME, szFile As String, szFileTitle As String 'Declare x for local use as long, ofn as openfilename, szfile as string, szfiletitle as string

    OFN.lStructSize = Len(OFN)
    OFN.hwnd = f.hwnd
    OFN.lpstrTitle = szDialogTitle
    OFN.lpstrFile = szFilename & String$(250 - Len(szFilename), 0)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = String$(255, 0)
    OFN.nMaxFileTitle = 255
    OFN.lpstrFilter = szFilter
    OFN.nFilterIndex = 1
    OFN.lpstrInitialDir = szDefDir
    OFN.lpstrDefExt = szDefExt

    If wMode = 1 Then
        OFN.Flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        X = GetOpenFileName(OFN)
    Else
        OFN.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
        X = GetSaveFileName(OFN)
    End If

    If X <> 0 Then

                                             '// If InStr(OFN.lpstrFileTitle, Chr$(0)) > 0 Then
                                             '//     szFileTitle = Left$(OFN.lpstrFileTitle, InStr(OFN.lpstrFileTitle, Chr$(0)) - 1)
                                             '// End If
        If InStr(OFN.lpstrFile, Chr$(0)) > 0 Then
            szFile = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, Chr$(0)) - 1)
        End If
                                             '// OFN.nFileOffset is the number of characters from the beginning of the
                                             '// full path to the start of the file name
                                             '// OFN.nFileExtension is the number of characters from the beginning of the
        '// full path to the file's extention, including the (.)
                                             '// MsgBox "File Name is " & szFileTitle & Chr$(13) & Chr$(10) & "Full path and file is " & szFile, , "Open"

                                             '// DialogFile = szFile & "|" & szFileTitle
        DialogFile = szFile

    Else

        DialogFile = ""                      'Empty dialogfile

    End If

End Function
Public Function DialogFont(f As Form, c As Control) As Boolean

    Dim LF As LOGFONT, FS As FONTSTRUC       'Declare lf for local use as logfont, fs as fontstruc
    Dim lLogFontAddress As Long, lMemHandle As Long 'Declare llogfontaddress for local use as long, lmemhandle as long

    If c.FontBold Then LF.lfWeight = FW_BOLD
    If c.FontItalic = True Then LF.lfItalic = 1
    If c.FontUnderline = True Then LF.lfUnderline = 1
    If c.FontStrikethru = True Then LF.lfStrikeOut = 1

    FS.lStructSize = Len(FS)

    lMemHandle = GlobalAlloc(GHND, Len(LF))
    If lMemHandle = 0 Then
        DialogFont = False
        Exit Function                        'Leave this function
    End If

    lLogFontAddress = GlobalLock(lMemHandle)
    If lLogFontAddress = 0 Then
        DialogFont = False
        Exit Function                        'Leave this function
    End If

    CopyMemory ByVal lLogFontAddress, LF, Len(LF)
    FS.lpLogFont = lLogFontAddress
    FS.iPointSize = c.FontSize * 10
    FS.Flags = CF_SCREENFONTS Or CF_EFFECTS

    If ChooseFont(FS) = 1 Then

        CopyMemory LF, ByVal lLogFontAddress, Len(LF)

        If LF.lfWeight >= FW_BOLD Then
            c.FontBold = True
        Else
            c.FontBold = False
        End If

        If LF.lfItalic = 1 Then
            c.FontItalic = True
        Else
            c.FontItalic = False
        End If

        If LF.lfUnderline = 1 Then
            c.FontUnderline = True
        Else
            c.FontUnderline = False
        End If

        If LF.lfStrikeOut = 1 Then
            c.FontStrikethru = True
        Else
            c.FontStrikethru = False
        End If

        c.FontName = ByteToString(LF.lfFaceName())
        c.FontSize = CLng(FS.iPointSize / 10)

        DialogFont = True

    Else

        DialogFont = False

    End If

End Function
Public Function DialogPrint(hwnd As Long, bPages As Boolean, Flags As Long) As PRINTPROPS

    Dim DM As DEVMODE, PD As PRINTDLGSTRUC   'Declare dm for local use as devmode, pd as printdlgstruc
    Dim lpDM As Long, wNull As Integer, szDevName As String 'Declare lpdm for local use as long, wnull as integer, szdevname as string

    PD.lStructSize = Len(PD)
    PD.hwnd = hwnd
    PD.hDevMode = 0                          'reset pd.hdevmode to zero
    PD.hDevNames = 0                         'reset pd.hdevnames to zero
    PD.hDC = 0                               'reset pd.hdc to zero
    PD.Flags = Flags
    PD.nFromPage = 0                         'reset pd.nfrompage to zero
    PD.nToPage = 0                           'reset pd.ntopage to zero
    PD.nMinPage = 0                          'reset pd.nminpage to zero
    If bPages Then PD.nMaxPage = bPages - 1
    PD.nCopies = 0                           'reset pd.ncopies to zero
    DialogPrint.Cancel = True

    If PrintDlg(PD) Then

        lpDM = GlobalLock(PD.hDevMode)
        CopyMemory DM, ByVal lpDM, Len(DM)
        lpDM = GlobalUnlock(PD.hDevMode)

        DialogPrint.Cancel = False

        DialogPrint.Device = Left$(DM.dmDeviceName, InStr(DM.dmDeviceName, Chr(0)) - 1)

        If PD.Flags And PD_PRINTTOFILE Then DialogPrint.ToFile = True Else DialogPrint.ToFile = False

        If PD.Flags And PD_PAGENUMS Then
            DialogPrint.Range = ppRangePages
            DialogPrint.FromPage = PD.nFromPage
            DialogPrint.ToPage = PD.nToPage
        ElseIf PD.Flags And PD_SELECTION Then
            DialogPrint.Range = ppRangeSelection
            DialogPrint.FromPage = 0         'reset dialogprint.frompage to zero
            DialogPrint.ToPage = 0           'reset dialogprint.topage to zero
        Else
            DialogPrint.Range = ppRangeAll
            DialogPrint.FromPage = 0         'reset dialogprint.frompage to zero
            DialogPrint.ToPage = 0           'reset dialogprint.topage to zero
        End If

        If PD.nCopies = 1 Then
            DialogPrint.Copies = DM.dmCopies
        End If

    End If

End Function
Public Function DialogPrintSetup(f As Form)

    Dim X As Long, PD As PRINTDLGSTRUC       'Declare x for local use as long, pd as printdlgstruc

    PD.lStructSize = Len(PD)
    PD.hwnd = f.hwnd
    PD.Flags = PD_PRINTSETUP
    X = PrintDlg(PD)

End Function

