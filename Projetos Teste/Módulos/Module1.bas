Attribute VB_Name = "Module1"
Public CN As ADODB.Connection
Public Rec As ADODB.Recordset

Public Function Abre_Conexao() As String

    Dim strSQL As String
    
    Set CN = New ADODB.Connection
        CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Mala\BDmala.mdb;Persist Security Info=False"

    SQL_Pesquisa = "Select * From TBclientes"
    
    Set Rec = New ADODB.Recordset
        Rec.CursorLocation = adUseClient
        Rec.Open SQL_Pesquisa, CN, adOpenKeyset, adLockOptimistic, adCmdText
        
End Function
