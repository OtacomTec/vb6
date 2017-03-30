VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   1635
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   390
      TabIndex        =   0
      Top             =   450
      Width           =   2025
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rela As GM_clRelatório
Private Sub Command1_Click()
    'Dim Rel As GM_clRelatório
    Set Rel = New GM_clRelatório
    
    Rel.CabeçalhoPágina "Nome Do Relatório", _
                        "15/10/2001", , , "SubTítulo"
                        
    Rel.FimDaImpressão
    
                       
End Sub
