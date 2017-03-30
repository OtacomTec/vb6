VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDados_Relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim str_Linha As String
Dim str_Negrito As String
Dim lng_IGrupo As Long

'Linha do Relatório
Public Property Let Linha(ByVal New_Linha As String)

    str_Linha = New_Linha

End Property

'Controle de Negrito
Public Property Let Negrito(ByVal New_Negrito As String)

    str_Negrito = New_Negrito

End Property

'Controle de Cabeçalho de Grupo
Public Property Let IGrupo(ByVal New_IGrupo As Long)

    lng_IGrupo = New_IGrupo

End Property

'Linha do Relatório
Public Property Get Linha() As String

    Linha = str_Linha

End Property

'Controle de Negrito
Public Property Get Negrito() As String

    Negrito = str_Negrito

End Property

'Controle de Cabeçalho de Grupo
Public Property Get IGrupo() As Long

    IGrupo = lng_IGrupo

End Property

